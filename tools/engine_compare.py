#!/usr/bin/env python3
"""Side-by-side engine comparison: Word reference | docxide-pdf | LibreOffice | MiniPdf | rdocx | office2pdf | jubarte-redlines | dxpdf.

Reuses PNGs the test harness already produced under tests/output/<group>/<case>/ (reference/, and
generated/ while its PDF is byte-identical to this run's conversion with the current binary). Conversions and
screenshots are cached in comparison/work/. Every engine is additionally timed by converting
into comparison/work/ once (seconds cached in a .time file beside the PDF). A conversion over
TIMEOUT (120 s) is killed, logged to comparison/work/timeouts.tsv and not retried until the document,
the engine, the fonts or the limit change; each run ends with the list. Scores are cached per case and engine
(<engine>.score.json), keyed by both PDFs' contents and a fingerprint of the scorer and the tools it
runs (page-metrics, mutool, veraPDF, pdfinfo, DPI); screenshots and accessibility analyses are redone
when those tools change. The viewer is a self-contained static site:
comparison/index.html plus lossless WebP page images, deployable as-is with
tools/deploy_comparison.sh (work/ is excluded by comparison/.gitignore).

Usage:
    python3 tools/engine_compare.py                 # every fixture with a reference.pdf
    python3 tools/engine_compare.py --case case41 --case 'case2*'   # exact or glob, repeatable
    python3 tools/engine_compare.py --group cases --open
    python3 tools/engine_compare.py --skip-libreoffice --no-scores
    python3 tools/engine_compare.py --fresh         # reconvert everything, ignoring every cache
    python3 tools/engine_compare.py --html-only     # rebuild index.html from the cached manifest, no re-scoring
    python3 tools/engine_compare.py --shard 1/4     # the 2nd of 4 shards (fixtures assigned by name hash): CI runs them in parallel ...
    python3 tools/engine_compare.py --merge comparison/manifest.*.json   # ... and builds the site from their manifests

rdocx: `rdocx` on PATH (cargo install rdocx) or RDOCX_BIN.
MiniPdf: the Rust crate's CLI, `minipdf` on PATH (cargo install minipdf-cli) or MINIPDF_BIN.
The .NET engine is a different implementation and is deliberately not what we compare against.
office2pdf: `office2pdf` on PATH (cargo install office2pdf-cli) or OFFICE2PDF_BIN.
jubarte-redlines: `jubarte` on PATH (cargo install jubarte-redlines) or JUBARTE_BIN.
dxpdf: `dxpdf` on PATH (cargo install dxpdf, or the .deb from its GitHub releases) or DXPDF_BIN.
Accessibility scores need verapdf and pdfinfo on PATH (brew install verapdf poppler); without them the column is empty.

Fonts, so every engine sees the same Word fonts: ours reads DOCXIDE_FONTS=fonts/; minipdf, rdocx,
office2pdf and jubarte get comparison/work/fonts_flat/, one directory of links to every file under fonts/
(rdocx takes a single directory and does not descend into fonts/CloudFonts/*; jubarte takes it as
JUBARTE_FONT_DIR and scans no Linux system directory at all, so without it CI would leave it with its
bundled Carlito/Liberation). LibreOffice has no font flag and reads
fontconfig / the OS: CI links fonts/ into ~/.local/share/fonts, which is the same files our discovery
already ranks first there, so nothing shifts. Do NOT do the same on macOS: a copy of fonts/ under
~/Library/Fonts outranks the macOS system faces our references were made with (Times New Roman,
Arial, Symbol) and collapses dozens of our fixtures. Locally, LibreOffice therefore only sees the fonts
the OS has; its CI column is the fair one. dxpdf is the same: no font flag, it reads fontconfig on Linux
and CoreText on macOS, so it gets fonts/ on CI only.
"""
from __future__ import annotations

import argparse
import filecmp
import functools
import hashlib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
import webbrowser
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from fnmatch import fnmatch
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = ROOT / "tests" / "fixtures"
TEST_OUTPUT = ROOT / "tests" / "output"
SITE = ROOT / "comparison"   # deployable: index.html + <group>/<case>/<engine>/page_NNN.webp
WORK = SITE / "work"         # cache: converted PDFs, PNG screenshots, LibreOffice profiles
MANIFEST = WORK / "manifest.json"
OURS_BIN = ROOT / "target" / "release" / "docxide-pdf"
METRICS_BIN = ROOT / "tools" / "target" / "release" / "page-metrics"
METRICS = ["jaccard", "ssim", "text_boundary"]
GROUPS = ["cases", "scraped"]
DPI = "150"  # same as tests/common/mod.rs MUTOOL_DPI
TIMEOUT = 120  # seconds per conversion, every engine; slower ones are logged to TIMEOUT_LOG
TIMEOUT_LOG = WORK / "timeouts.tsv"
TIMEOUTS: list[tuple[str, str, bool]] = []  # (engine, group/case, timed out this run) for the summary

# (key, label). Key doubles as the PNG directory name.
ENGINES = [
    ("reference", "Word"),
    ("generated", "docxide-pdf"),
    ("libreoffice", "LibreOffice"),
    ("minipdf", "MiniPdf (Rust)"),
    ("rdocx", "rdocx"),
    ("office2pdf", "office2pdf"),
    ("jubarte", "jubarte-redlines"),
    ("dxpdf", "dxpdf"),
]
COMPETITORS = [k for k, _ in ENGINES if k != "reference"]

# Scores are cached per case and engine, keyed by content rather than mtimes (CI dates files by
# commit, the harness rewrites its PDFs). SCORING_FP covers everything besides the two PDFs that a
# score depends on; main() sets it, and an empty one means "do not cache".
SCORING_FP = ""
A11Y_FP_FILE = WORK / "a11y.fingerprint"
A11Y_SRC = ROOT / "tests" / "common" / "a11y.rs"


# Seconds spent per phase, summed over the worker threads, so a CI log says where a run's time went.
SPENT: Counter = Counter()
_SPENT_LOCK = threading.Lock()
# (phase, reason) → engine-cases this run rendered or scored again, for the summary line.
REDONE: Counter = Counter()


def spent(phase: str, fn, *args, **kwargs):
    t = time.perf_counter()
    try:
        return fn(*args, **kwargs)
    finally:
        with _SPENT_LOCK:
            SPENT[phase] += time.perf_counter() - t


def report_timeouts() -> None:
    """This run's conversions over the limit: timed out now, or skipped as a known timeout."""
    if not TIMEOUTS:
        return
    print(f"{len(TIMEOUTS)} conversions over the {TIMEOUT} s limit, not retried until the document, engine, "
          f"fonts or limit change (--fresh retries them; history in {TIMEOUT_LOG.relative_to(ROOT)}):")
    for engine, case, new in sorted(TIMEOUTS, key=lambda t: (t[1], t[0])):
        print(f"  {engine:12} {case}{'' if new else '  (known, skipped)'}")


def report_spent() -> None:
    if SPENT:
        print("time spent (thread-seconds): " + ", ".join(f"{k} {v:.0f}" for k, v in SPENT.most_common()))


def report_redone() -> None:
    """What this run rendered or scored again, and why: a slow run should be readable from its log."""
    by_phase: dict[str, Counter] = {}
    for (phase, reason), n in REDONE.items():
        by_phase.setdefault(phase, Counter())[reason] += n
    if by_phase:
        print("redone: " + "; ".join(f"{phase} {sum(c.values())} ({', '.join(f'{r} {n}' for r, n in c.most_common())})"
                                     for phase, c in by_phase.items()))


def find_soffice() -> Path | None:
    env = os.environ.get("LIBREOFFICE_PATH")
    if env and Path(env).is_file():
        return Path(env)
    mac = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if mac.is_file():
        return mac
    found = shutil.which("soffice")
    return Path(found) if found else None


def find_minipdf() -> Path | None:
    """The Rust crate's CLI (cargo install minipdf-cli), not the .NET engine's native binary."""
    env = os.environ.get("MINIPDF_BIN")
    if env and Path(env).is_file():
        return Path(env)
    found = shutil.which("minipdf") or str(Path.home() / ".cargo" / "bin" / "minipdf")
    return Path(found) if Path(found).is_file() else None


def find_rdocx() -> Path | None:
    env = os.environ.get("RDOCX_BIN")
    if env and Path(env).is_file():
        return Path(env)
    found = shutil.which("rdocx")
    return Path(found) if found else None


def find_office2pdf() -> Path | None:
    env = os.environ.get("OFFICE2PDF_BIN")
    if env and Path(env).is_file():
        return Path(env)
    found = shutil.which("office2pdf") or str(Path.home() / ".cargo" / "bin" / "office2pdf")
    return Path(found) if Path(found).is_file() else None


def find_jubarte() -> Path | None:
    env = os.environ.get("JUBARTE_BIN")
    if env and Path(env).is_file():
        return Path(env)
    found = shutil.which("jubarte") or str(Path.home() / ".cargo" / "bin" / "jubarte")
    return Path(found) if Path(found).is_file() else None


def find_dxpdf() -> Path | None:
    env = os.environ.get("DXPDF_BIN")
    if env and Path(env).is_file():
        return Path(env)
    found = shutil.which("dxpdf") or str(Path.home() / ".cargo" / "bin" / "dxpdf")
    return Path(found) if Path(found).is_file() else None


def ensure_built(bin_path: Path, cwd: Path, *cargo_args: str, always: bool = False) -> Path | None:
    if bin_path.is_file() and not always:
        return bin_path
    print(f"building {bin_path.name} ...")
    r = subprocess.run(["cargo", "build", *cargo_args], cwd=cwd, capture_output=True, text=True)
    if r.returncode != 0:
        print(r.stderr[-800:], file=sys.stderr)
        return None
    return bin_path if bin_path.is_file() else None


def is_fresh(target: Path, source: Path) -> bool:
    return target.exists() and target.stat().st_mtime >= source.stat().st_mtime


def screenshot(pdf: Path, out_dir: Path) -> tuple[list[Path], str | None]:
    """Render every page to out_dir/page_NNN.png, unless the PNGs are newer than the PDF and
    RENDER_EPOCH. Returns the pages and why they were rendered (None: reused)."""
    existing = sorted(out_dir.glob("page_*.png"))
    pdf_time = pdf.stat().st_mtime
    if existing:
        oldest = min(p.stat().st_mtime for p in existing)
        if oldest >= max(pdf_time, RENDER_EPOCH):
            return existing, None
        why = "pdf changed" if oldest < pdf_time else "render tools changed"
    else:
        why = "new"
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in existing:
        old.unlink()
    subprocess.run(
        ["mutool", "draw", "-F", "png", "-r", DPI, "-o", str(out_dir / "page_%03d.png"), str(pdf)],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False,
    )
    return sorted(out_dir.glob("page_*.png")), why


def converted(r: subprocess.CompletedProcess, pdf: Path) -> bool:
    """Engines fail quietly on some documents; print the tail of their output so a CI log says why."""
    if pdf.exists():
        return True
    why = " | ".join(((r.stderr or "") + (r.stdout or "")).strip().splitlines()[-3:])[-300:]
    print(f"  {Path(r.args[0]).name} failed on {pdf.parent.name} (exit {r.returncode}): {why}")
    return False


def convert_ours(docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx) and is_fresh(pdf, OURS_BIN):  # a rebuilt binary must be re-timed
        return True
    pdf.parent.mkdir(parents=True, exist_ok=True)
    # The binary reads DOCXIDE_FONTS at run time; `.cargo/config.toml` only sets it under cargo,
    # so without this a local run renders with system fonts while CI renders with the Word fonts.
    env = {**os.environ}
    env.setdefault("DOCXIDE_FONTS", str(ROOT / "fonts"))
    # Convert to a fresh path (the CLI never overwrites; it would write generated(2).pdf next to a
    # cached PDF and leave the stale one in place) and keep the cached PDF when the output is
    # byte-identical, which it is for every case a commit does not touch, conversion being
    # deterministic. Its mtime then stays put, so the screenshots, site images and accessibility
    # analysis derived from it stay fresh and a warm run only redoes the cases that changed.
    tmp = pdf.with_name(pdf.stem + ".tmp.pdf")
    tmp.unlink(missing_ok=True)
    t = time.perf_counter()
    r = subprocess.run([str(OURS_BIN), str(docx), str(tmp)], capture_output=True, text=True, env=env, timeout=TIMEOUT)
    seconds = time.perf_counter() - t
    if r.returncode != 0 or not tmp.exists():
        tmp.unlink(missing_ok=True)
        pdf.unlink(missing_ok=True)  # a stale PDF must not stand in for a failed conversion
        return converted(r, pdf)
    if pdf.exists() and pdf.read_bytes() == tmp.read_bytes():
        tmp.unlink()
        pdf.with_suffix(".time").write_text(f"{seconds:.3f}")  # timed() only re-times a rewritten file
    else:
        tmp.replace(pdf)
    return True


def convert_libreoffice(soffice: Path, docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx) and is_fresh(pdf, soffice):  # an upgraded engine reconverts
        return True
    out = pdf.parent
    out.mkdir(parents=True, exist_ok=True)
    # Per-case profile dir sidesteps LibreOffice's single-instance lock (same trick as the harness).
    profile = (out / "lo_profile").resolve()
    profile.mkdir(exist_ok=True)
    r = subprocess.run(
        [str(soffice), f"-env:UserInstallation=file://{profile}", "--headless",
         "--convert-to", "pdf", "--outdir", str(out), str(docx)],
        capture_output=True, text=True, timeout=TIMEOUT, check=False,
    )
    produced = out / (docx.stem + ".pdf")
    if produced.exists() and produced != pdf:
        produced.replace(pdf)
    shutil.rmtree(profile, ignore_errors=True)   # throwaway, and CI would otherwise cache it forever
    return converted(r, pdf)


def font_face_name(path: Path) -> str | None:
    """'<family>-<subfamily>' from the font's own name table (first face of a collection), or None.
    Stdlib only: CI has no fontTools."""
    import struct
    try:
        data = path.read_bytes()
        off = 0
        if data[:4] == b"ttcf":
            off = struct.unpack_from(">I", data, 12)[0]
        num_tables = struct.unpack_from(">H", data, off + 4)[0]
        for i in range(num_tables):
            tag, _, toff, _ = struct.unpack_from(">4sIII", data, off + 12 + 16 * i)
            if tag == b"name":
                break
        else:
            return None
        count, strings = struct.unpack_from(">HH", data, toff + 2)
        names: dict[int, str] = {}
        for i in range(count):
            plat, enc, lang, nid, length, soff = struct.unpack_from(">HHHHHH", data, toff + 6 + 12 * i)
            if plat == 3 and nid in (1, 2, 16, 17) and (nid not in names or lang == 0x409):
                raw = data[toff + strings + soff: toff + strings + soff + length]
                names[nid] = raw.decode("utf-16-be", "replace")
        family, sub = names.get(16) or names.get(1), names.get(17) or names.get(2) or "Regular"
        if not family:
            return None
        clean = lambda s: re.sub(r"[^A-Za-z0-9 ]+", "", s).strip()
        return f"{clean(family)}-{clean(sub)}"
    except (OSError, struct.error, IndexError):
        return None


_FLAT_FONT_LOCK = threading.Lock()


def flat_font_dir() -> Path | None:
    with _FLAT_FONT_LOCK:   # cases convert in parallel; the first caller builds, the rest wait for it
        return _flat_font_dir()


@functools.lru_cache(maxsize=None)
def _flat_font_dir() -> Path | None:
    """One directory of links to every font file under fonts/, for engines that take a single font
    directory and do not descend into fonts/CloudFonts/<family>/ (rdocx, jubarte). Rebuilt per run.
    Word's cloud-font cache names files by number (fonts/CloudFonts/Vivaldi/19672202630.ttf) and jubarte
    finds a family by file name before it confirms it against the name table, so those links are named
    '<family>-<subfamily>' from the font itself; a name that repeats is prefixed with its folder."""
    src = ROOT / "fonts"
    if not src.is_dir():
        return None
    flat = WORK / "fonts_flat"
    shutil.rmtree(flat, ignore_errors=True)
    flat.mkdir(parents=True)
    for f in sorted(p for p in src.rglob("*") if p.suffix.lower() in (".ttf", ".ttc", ".otf")):
        name = f.name
        if f.stem.isdigit():
            face = font_face_name(f)
            if face:
                name = face + f.suffix.lower()
        dst = flat / name
        if dst.exists() or dst.is_symlink():
            dst = flat / f"{f.parent.name}__{name}"
        dst.symlink_to(f.resolve())
    return flat


def convert_minipdf(minipdf: Path, docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx) and is_fresh(pdf, minipdf):
        return True
    pdf.parent.mkdir(parents=True, exist_ok=True)
    # Same Word fonts as the other engines. Without --fonts, minipdf 0.6 registers a hard-coded list of
    # Linux system fonts and panics ("UnknownKind") on most documents.
    fonts = ["--fonts", str(flat_font_dir())] if flat_font_dir() else []
    r = subprocess.run([str(minipdf), "convert", str(docx), "-o", str(pdf), *fonts],
                       capture_output=True, text=True, timeout=TIMEOUT, check=False)
    return converted(r, pdf)


def convert_rdocx(rdocx: Path, docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx) and is_fresh(pdf, rdocx):
        return True
    pdf.parent.mkdir(parents=True, exist_ok=True)
    # Same Word fonts as the other engines; without --font-dir it draws Calibri documents in its bundled Carlito.
    fonts = ["--font-dir", str(flat_font_dir())] if flat_font_dir() else []
    r = subprocess.run([str(rdocx), "convert", "--to", "pdf", "--output", str(pdf), *fonts, str(docx)],
                       capture_output=True, text=True, timeout=TIMEOUT, check=False)
    return converted(r, pdf)


def convert_office2pdf(office2pdf: Path, docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx) and is_fresh(pdf, office2pdf):
        return True
    pdf.parent.mkdir(parents=True, exist_ok=True)
    # Same Word fonts as the other engines; otherwise it uses whatever the host happens to have installed.
    fonts = ["--font-path", str(flat_font_dir())] if flat_font_dir() else []
    r = subprocess.run([str(office2pdf), str(docx), "-o", str(pdf), *fonts],
                       capture_output=True, text=True, timeout=TIMEOUT, check=False)
    return converted(r, pdf)


def convert_jubarte(jubarte: Path, docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx) and is_fresh(pdf, jubarte):
        return True
    pdf.parent.mkdir(parents=True, exist_ok=True)
    # Its own font folder; the only font directory it reads on Linux (macOS adds Word's DFonts and the
    # system folders by itself). Without it, CI renders with its bundled Carlito/Liberation.
    env = {**os.environ}
    if flat_font_dir():
        env["JUBARTE_FONT_DIR"] = str(flat_font_dir())
        env["JUBARTE_FONT_INDEX"] = "off"   # its on-disk index is keyed by the folder's path, not its contents
    r = subprocess.run([str(jubarte), "convert", str(docx), "-o", str(pdf), "--force"],
                       capture_output=True, text=True, env=env, timeout=TIMEOUT, check=False)
    return converted(r, pdf)


def convert_dxpdf(dxpdf: Path, docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx) and is_fresh(pdf, dxpdf):
        return True
    pdf.parent.mkdir(parents=True, exist_ok=True)
    # No font flag: it sees fonts/ only through fontconfig, which CI points at it (see the module docstring).
    r = subprocess.run([str(dxpdf), str(docx), "-o", str(pdf)],
                       capture_output=True, text=True, timeout=TIMEOUT, check=False)
    return converted(r, pdf)


@functools.lru_cache(maxsize=None)
def engine_id(engine: Path) -> str:
    """What identifies an engine build: our binary's bytes (rebuilt from the working tree every
    run), a competitor's version banner (a reinstall keeps it, an upgrade changes it)."""
    return file_hash(engine) if engine == OURS_BIN else tool_version(str(engine), "--version")


@functools.lru_cache(maxsize=None)
def fonts_fingerprint() -> str:
    """Every engine converts with fonts/, so adding or replacing a font can change what hangs."""
    d = ROOT / "fonts"
    if not d.is_dir():
        return "no fonts"
    return fingerprint(*(f"{p.relative_to(d)}:{p.stat().st_size}" for p in sorted(d.rglob("*")) if p.is_file()))


def timed(convert, *args) -> tuple[bool, float | None]:
    """Run a convert_* and return (ok, wall-clock seconds). Seconds are stored beside the PDF so a
    cached conversion keeps its measured time; a cached PDF without one is converted again.
    # ponytail: measured under --jobs parallel conversions; use --jobs 1 for clean absolute numbers.
    """
    pdf: Path = args[-1]
    docx: Path = args[-2]
    engine: Path = args[0] if len(args) == 3 else OURS_BIN  # convert_ours takes no binary argument
    stamp = pdf.with_suffix(".time")
    if pdf.exists() and not stamp.exists():
        pdf.unlink()
    # A timeout leaves no PDF to cache, so without this marker the engine would hang for the full
    # timeout on the same document every run (minipdf: 2 × 300 s per run). Retried once the
    # document, the engine, the fonts or the limit change: the marker holds a key of all four.
    timeout_marker = pdf.with_suffix(".timeout")
    case = pdf.parent.relative_to(WORK).as_posix()
    timeout_key = fingerprint(str(TIMEOUT), file_hash(docx), engine_id(engine), fonts_fingerprint())
    if timeout_marker.exists() and timeout_marker.read_text() == timeout_key:
        with _SPENT_LOCK:
            TIMEOUTS.append((pdf.stem, case, False))
        return False, None
    # A competitor's cached PDF is dropped when its engine changed version. The binary's mtime alone
    # does not catch that: apt dates soffice by the package's build, older than any cached PDF. (Ours
    # is rebuilt every run; convert_ours reconverts and keeps the PDF when the bytes are unchanged.)
    engine_stamp = pdf.with_suffix(".engine")
    if engine != OURS_BIN and pdf.exists() and (not engine_stamp.exists() or engine_stamp.read_text() != engine_id(engine)):
        pdf.unlink()
    before = pdf.stat().st_mtime if pdf.exists() else None
    t = time.perf_counter()
    try:
        ok = convert(*args)
    except subprocess.TimeoutExpired as e:  # one hung engine must not drop the whole case
        print(f"  {pdf.stem} timed out after {e.timeout:.0f} s on {case}")
        pdf.parent.mkdir(parents=True, exist_ok=True)
        timeout_marker.write_text(timeout_key)
        with _SPENT_LOCK, TIMEOUT_LOG.open("a") as log:
            log.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}\t{pdf.stem}\t{case}\t{e.timeout:.0f}\n")
            TIMEOUTS.append((pdf.stem, case, True))
        return False, None
    if ok:
        timeout_marker.unlink(missing_ok=True)  # converts in time again (new engine or document)
        if engine != OURS_BIN:
            engine_stamp.write_text(engine_id(engine))
    if ok and pdf.stat().st_mtime != before:
        stamp.write_text(f"{time.perf_counter() - t:.3f}")
    return ok, float(stamp.read_text()) if ok and stamp.exists() else None


def run_out(cmd: list[str]) -> str:
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=30).stdout.strip()
    except (OSError, subprocess.TimeoutExpired):
        return ""


# Word does not write its version into the PDF (Creator is just "Microsoft Word"), so this is
# recorded by hand. Keep in sync with the "Reference PDFs are generated using ..." note in README.md.
WORD_VERSION = "Word for Mac 16.106.1, online export"


def engine_versions(tools: dict) -> dict[str, str]:
    """Version string per engine, captured at run time so the site says what produced its images."""
    v: dict[str, str] = {"reference": WORD_VERSION}
    m = re.search(r'^version\s*=\s*"([^"]+)"', (ROOT / "Cargo.toml").read_text(), re.M)
    sha = run_out(["git", "-C", str(ROOT), "rev-parse", "--short", "HEAD"])
    dirty = "+dirty" if run_out(["git", "-C", str(ROOT), "status", "--porcelain", "--", "src", "Cargo.toml"]) else ""
    v["generated"] = f"{m.group(1) if m else '?'} @{sha}{dirty}"
    if tools.get("soffice"):
        v["libreoffice"] = " ".join(run_out([str(tools["soffice"]), "--version"]).split()[:2])  # drop the build hash
    for key in ("minipdf", "rdocx", "office2pdf", "jubarte", "dxpdf"):
        if tools.get(key):
            v[key] = run_out([str(tools[key]), "--version"]).split()[-1]
    return v


def pdf_creator(pdf: Path) -> str:
    """Creator (or Producer) from the PDF Info dict; Word writes 'Microsoft Word' with no version."""
    info = run_out(["mutool", "info", str(pdf)])
    for tag in ("Creator", "Producer"):
        m = re.search(r"/" + tag + r"\(([^)]*)\)", info)
        if m:
            return m.group(1)
    return ""


def engine_metrics(ref_pdf: Path, other_pdf: Path, ref_dir: Path, other_dir: Path) -> dict:
    """Jaccard, SSIM and text-boundary in percent, plus the accessibility scores (rule counts as is,
    structure/text in percent), computed by tools/page-metrics with the harness's own code."""
    if not METRICS_BIN.is_file():
        return {}
    r = subprocess.run([str(METRICS_BIN), str(ref_pdf), str(other_pdf), str(ref_dir), str(other_dir)],
                       capture_output=True, text=True)
    try:
        m = json.loads(r.stdout)
    except ValueError:
        return {}
    out = {k: round(m[k] * 100, 1) for k in METRICS if m.get(k) is not None}
    if m.get("a11y"):
        out["a11y"] = {k: round(v * 100, 1) if isinstance(v, float) else v for k, v in m["a11y"].items()}
    return out


def file_hash(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def tool_version(*cmd: str) -> str:
    """A tool's version banner (mutool and pdfinfo print it on stderr), or "missing"."""
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        return (r.stdout + r.stderr).strip()
    except (OSError, subprocess.TimeoutExpired):
        return "missing"


def fingerprint(*parts: str) -> str:
    return hashlib.sha256("\n".join(parts).encode()).hexdigest()


def record_fingerprint(marker: Path, fp: str) -> bool:
    """Store fp in marker; True when it differs from the one stored by an earlier run. The first
    run has nothing to compare with and trusts the caches it finds (--fresh redoes a case anyway)."""
    known = marker.read_text() if marker.exists() else None
    if known == fp:
        return False
    WORK.mkdir(parents=True, exist_ok=True)
    marker.write_text(fp)
    if known is None:
        os.utime(marker, (0, 0))  # an epoch-dated marker makes no screenshot look older than it
    return known is not None


RENDER_FP_FILE = WORK / "render.fingerprint"
RENDER_EPOCH = 0.0  # screenshots older than this were rendered by another mutool or DPI


def render_check() -> None:
    """Screenshots are only re-rendered when older than their PDF, so a mutool upgrade or DPI change
    moves RENDER_EPOCH forward instead: everything rendered before it, the harness's renders that
    are reused here included, is redone in place (the site images and visual scores follow)."""
    global RENDER_EPOCH
    if record_fingerprint(RENDER_FP_FILE, fingerprint(tool_version("mutool", "-v"), DPI)):
        print("mutool or DPI changed: every screenshot is rendered again")
    RENDER_EPOCH = RENDER_FP_FILE.stat().st_mtime


def scoring_fingerprint() -> str:
    """The scorer and every tool it shells out to. page-metrics reuses the accessibility analyses
    (*.a11y.json) while they are newer than the PDF, so those go when their own inputs change."""
    a11y = fingerprint(file_hash(A11Y_SRC), tool_version("verapdf", "--version"), tool_version("pdfinfo", "-v"))
    if record_fingerprint(A11Y_FP_FILE, a11y):
        # The harness's analyses too: it reuses them on the same mtime test, so they are stale for it as well.
        stale = [*WORK.glob("*/*/*.a11y.json"), *TEST_OUTPUT.glob("*/*/*.a11y.json")]
        for f in stale:
            f.unlink(missing_ok=True)
        print(f"veraPDF, pdfinfo or {A11Y_SRC.name} changed: dropped {len(stale)} cached analyses")
    # The scorer by its sources, not its binary: page-metrics links the whole library (tests/common
    # calls the converter), so the binary changes with nearly every commit while scoring does not,
    # and CI recompiles it every run (rust-cache keeps dependencies only).
    sources = [ROOT / "tools" / "src" / "bin" / "page_metrics.rs", *sorted((ROOT / "tests" / "common").glob("*.rs"))]
    # Of the lock file, only the crates the scorer runs on: a bump of the case browser's GUI crate, or
    # of rustc, changes no score, and either used to cost a full rescore (~25 min per CI shard).
    lock = (ROOT / "tools" / "Cargo.lock").read_text()
    crates = re.findall(r'^name = "(image|png|rayon|roxmltree|serde_json|sha2|zip)"\nversion = "([^"]+)"', lock, re.M)
    scorer = [file_hash(p) for p in sources] + [f"{name} {version}" for name, version in crates]
    return fingerprint(*scorer, DPI, tool_version("mutool", "-v"), a11y)


def cached_metrics(engine: str, mine: Path, ref_pdf: Path, ref_hash: str, pdf: Path,
                   ref_dir: Path, other_dir: Path) -> tuple[dict, str | None]:
    """engine_metrics, reused while both PDFs and the scoring fingerprint are unchanged. Returns the
    scores and why they were computed (None: reused)."""
    pdf_hash = file_hash(pdf)
    key = f"{SCORING_FP}:{ref_hash}:{pdf_hash}" if SCORING_FP else None
    cache = mine / f"{engine}.score.json"
    why = "new"
    if key:
        try:
            c = json.loads(cache.read_text())
            if c["key"] == key:
                return c["scores"], None
            old = c["key"].split(":")
            why = ("scorer changed" if old[0] != SCORING_FP else "reference changed" if old[1] != ref_hash
                   else "pdf changed")
        except (OSError, ValueError, KeyError, IndexError):
            pass
    m = spent("score", engine_metrics, ref_pdf, pdf, ref_dir, other_dir)
    if m and key:  # an empty result is a failed run of page-metrics, not a score
        mine.mkdir(parents=True, exist_ok=True)
        cache.write_text(json.dumps({"key": key, "scores": m}))
    return m, why


def process_fixture(fixture: Path, group: str, tools: dict, opts) -> dict | None:
    docx = fixture / "input.docx"
    ref_pdf = fixture / "reference.pdf"
    if not (docx.exists() and ref_pdf.exists()):
        return None
    case = fixture.name
    harness = TEST_OUTPUT / group / case
    mine = WORK / group / case
    pages: dict[str, list[Path]] = {}
    pdfs: dict[str, Path] = {}
    redone: list[tuple[str, str, str]] = []   # (phase, engine, reason): what this case had done again

    def add(key: str, pdf: Path, png_dir: Path) -> None:
        pdfs[key] = pdf
        pages[key], why = spent("screenshot", screenshot, pdf, png_dir)
        if why:
            redone.append(("render", key, why))

    # Word reference
    ref_dir = harness / "reference" if any((harness / "reference").glob("page_*.png")) else mine / "reference"
    add("reference", ref_pdf, ref_dir)

    # --fresh: drop cached conversions, timeout markers and scores so every engine runs again.
    if opts.fresh:
        for pattern in ("*.pdf", "*.time", "*.timeout", "*.score.json", "*.a11y.json"):
            for f in mine.glob(pattern):
                f.unlink()

    def show_conversion(key: str) -> None:
        """Show the conversion this run made with the current binary. The harness's copy (and its
        renders) stands in only when byte-identical: it is as old as the last run-tests.sh, so after
        a source change it shows the old code. LibreOffice stamps a creation date, so it never does."""
        own, theirs = mine / f"{key}.pdf", harness / f"{key}.pdf"
        if not opts.fresh and theirs.exists() and filecmp.cmp(theirs, own, shallow=False):
            add(key, theirs, harness / key)
        else:
            add(key, own, mine / key)

    times: dict[str, float | None] = {}
    if tools.get("ours"):
        ok, times["generated"] = spent("convert generated", timed, convert_ours, docx, mine / "generated.pdf")
        if ok:
            show_conversion("generated")

    if tools.get("soffice"):
        ok, times["libreoffice"] = spent("convert libreoffice", timed, convert_libreoffice, tools["soffice"], docx,
                                         mine / "libreoffice.pdf")
        if ok:
            show_conversion("libreoffice")

    for key, convert in (("minipdf", convert_minipdf), ("rdocx", convert_rdocx),
                         ("office2pdf", convert_office2pdf), ("jubarte", convert_jubarte),
                         ("dxpdf", convert_dxpdf)):
        if tools.get(key):
            ok, times[key] = spent(f"convert {key}", timed, convert, tools[key], docx, mine / f"{key}.pdf")
            if ok:
                add(key, mine / f"{key}.pdf", mine / key)

    scores: dict[str, dict] = {}
    if not opts.no_scores:
        ref_hash = file_hash(ref_pdf)
        for key in COMPETITORS:
            if pages.get(key):
                m, why = cached_metrics(key, mine, ref_pdf, ref_hash, pdfs[key], ref_dir, pages[key][0].parent)
                if m:
                    scores[key] = m
                    if why:
                        redone.append(("score", key, why))

    with _SPENT_LOCK:
        for phase, _, why in redone:
            REDONE[(phase, why)] += 1
    rel = lambda p: os.path.relpath(p, WORK)  # noqa: E731
    return {
        "group": group,
        "case": case,
        "pages": {k: [rel(p) for p in v] for k, v in pages.items()},
        "scores": scores,
        "times": {k: v for k, v in times.items() if v is not None},
        "reference_app": pdf_creator(ref_pdf),
        "redone": redone,   # for the log line only; main() drops it before the manifest
    }


# The page itself (HTML, CSS and script) lives next to this file; write_html() fills in its placeholders.
PAGE_TEMPLATE = Path(__file__).with_name("engine_compare.html")


def natural_key(s: str) -> list:
    """case1 < case2 < case10, not case1 < case10 < case2."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def shard_of(group: str, case: str, n: int) -> int:
    """Which of n shards a fixture belongs to. By name hash, not round robin over the sorted list: a
    new fixture then moves no other one to a shard whose cache has to convert and score it from scratch."""
    # ponytail: unweighted; balance by page count if one shard keeps lagging
    return int(hashlib.sha256(f"{group}/{case}".encode()).hexdigest(), 16) % n


def prune_work(fixtures: list[tuple[str, Path]]) -> None:
    """Drop cached conversions and site pages of fixtures this run does not own: deleted or renamed ones,
    and in a shard run the other shards' (CI caches comparison/ per shard, and nothing else would ever
    empty it; a stale site page would even be deployed)."""
    keep = {(g, d.name) for g, d in fixtures}
    gone = [d for base in (WORK, SITE) for g in GROUPS if (base / g).is_dir() for d in (base / g).iterdir()
            if d.is_dir() and (g, d.name) not in keep]
    for d in gone:
        shutil.rmtree(d, ignore_errors=True)
    if gone:
        print(f"pruned {len(gone)} case directories this run does not own")


def build_site(results: list[dict], versions: dict, fmt: str, jobs: int) -> None:
    """Deployable static site at comparison/: index.html + every page image under <group>/<case>/<engine>/.

    Images are re-encoded (not linked) so the site is a snapshot that survives later test runs.
    fmt="webp" is lossless via cwebp: pixel-identical and ~3.5x smaller than the mutool PNGs.
    (JPEG and lossy WebP were measured *larger* than PNG on these mostly-white pages.)
    """
    if fmt == "webp" and not shutil.which("cwebp"):
        sys.exit("--format webp needs cwebp (brew install webp); or use --format png")
    jobs_list: list[tuple[Path, Path]] = []
    rewritten: list[dict] = []
    for c in results:
        pages = {}
        for eng, files in c["pages"].items():
            new = []
            for rel in files:
                src = WORK / rel
                dst = SITE / c["group"] / c["case"] / eng / (Path(rel).stem + "." + fmt)
                src_time, dst_time = mtime(src), mtime(dst)
                if src_time is None:
                    # A filtered run carries unprocessed cases over from the manifest; their PNGs may be
                    # gone (a later test run rewrote tests/output). Keep the image already in the site.
                    if dst_time is None:
                        continue
                elif dst_time is None or dst_time < src_time:
                    jobs_list.append((src, dst))
                new.append(dst.relative_to(SITE).as_posix())
            pages[eng] = new
        rewritten.append({**c, "pages": pages})

    def transfer(pair: tuple[Path, Path]) -> None:
        src, dst = pair
        dst.parent.mkdir(parents=True, exist_ok=True)
        if fmt == "png":
            shutil.copy2(src, dst)
        else:
            spent("encode", subprocess.run, ["cwebp", "-quiet", "-lossless", str(src), "-o", str(dst)],
                  stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)

    print(f"site: {len(jobs_list)} images to encode into {SITE}")
    with ThreadPoolExecutor(max_workers=jobs) as pool:
        list(pool.map(transfer, jobs_list))
    write_html(rewritten, versions, SITE / "index.html")
    (SITE / ".nojekyll").touch()  # GitHub Pages: serve as-is, no Jekyll pass over 9k files
    (SITE / ".gitignore").write_text("/work/\n")  # deploy_comparison.sh commits this folder; keep the cache out
    total = 0
    for d, dirs, files in os.walk(SITE):
        dirs[:] = [x for x in dirs if Path(d, x) != WORK]   # the conversion cache is not part of the site
        total += sum(os.path.getsize(os.path.join(d, f)) for f in files)
    print(f"site ready: {SITE / 'index.html'} ({total / 1e6:.0f} MB)")


def mtime(path: Path) -> float | None:
    try:
        return path.stat().st_mtime
    except FileNotFoundError:
        return None


def write_html(results: list[dict], versions: dict, out: Path) -> None:
    page = (PAGE_TEMPLATE.read_text()
            .replace("__ENGINES__", json.dumps(ENGINES))
            .replace("__NENGINES__", str(len(ENGINES)))
            .replace("__METRICS__", json.dumps(METRICS))
            .replace("__VERSIONS__", json.dumps(versions))
            .replace("__DATA__", json.dumps(results, separators=(",", ":"))))   # last: the large one
    out.write_text(page)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--case", action="append", default=[], help="case name or glob, e.g. case41 or 'case4*' (repeatable)")
    ap.add_argument("--group", choices=GROUPS, action="append", default=[], help="fixture group (repeatable)")
    ap.add_argument("--skip-libreoffice", action="store_true")
    ap.add_argument("--skip-minipdf", action="store_true")
    ap.add_argument("--skip-rdocx", action="store_true")
    ap.add_argument("--skip-office2pdf", action="store_true")
    ap.add_argument("--skip-jubarte", action="store_true")
    ap.add_argument("--skip-dxpdf", action="store_true")
    ap.add_argument("--no-scores", action="store_true", help="skip Jaccard scoring")
    ap.add_argument("--fresh", action="store_true",
                    help="reconvert with every engine, ignoring cached PDFs, timeouts and the test harness's PDFs")
    ap.add_argument("--jobs", type=int, default=os.cpu_count() or 4)
    ap.add_argument("--open", action="store_true", help="open the HTML when done")
    ap.add_argument("--html-only", action="store_true",
                    help="rebuild comparison/index.html from comparison/work/manifest.json (no conversion or scoring)")
    ap.add_argument("--shard", metavar="I/N",
                    help="only the I-th of N shards (0-based); fixtures are assigned by name hash, and the cache is pruned to the shard")
    ap.add_argument("--merge", nargs="+", metavar="MANIFEST",
                    help="build the site from these shard manifests (no conversion or scoring)")
    ap.add_argument("--format", choices=["webp", "png"], default="webp",
                    help="site image format; webp is lossless and ~3.5x smaller than png (needs cwebp)")
    opts = ap.parse_args()

    def load_manifest() -> dict:
        m = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else []
        return {"versions": {}, "cases": m} if isinstance(m, list) else m  # pre-versions manifests were a bare list

    def finish(results: list[dict], versions: dict) -> None:
        # A filtered run (--case/--group) updates just those entries; the site keeps every other case.
        if opts.case or opts.group:
            old = load_manifest()
            done = {(c["group"], c["case"]) for c in results}
            results = [c for c in old["cases"] if c["group"] in GROUPS and (c["group"], c["case"]) not in done] + results
            versions = {**old["versions"], **versions}
        results.sort(key=lambda r: (GROUPS.index(r["group"]), natural_key(r["case"])))
        WORK.mkdir(parents=True, exist_ok=True)
        MANIFEST.write_text(json.dumps({"versions": versions, "cases": results}))  # --html-only rebuilds from this
        print(f"{len(results)} cases; engines: " + ", ".join(f"{k} {v}" for k, v in versions.items()))
        build_site(results, versions, opts.format, opts.jobs)
        report_spent()
        if opts.open:
            webbrowser.open((SITE / "index.html").as_uri())

    if opts.html_only:
        m = load_manifest()
        finish(m["cases"], m["versions"])
        return
    if opts.merge:
        parts = [json.loads(Path(p).read_text()) for p in opts.merge]
        finish([c for m in parts for c in m["cases"]], {k: v for m in parts for k, v in m["versions"].items()})
        return

    if not shutil.which("mutool"):
        sys.exit("mutool not found (brew install mupdf-tools)")
    render_check()
    tools: dict = {}
    # Always rebuild: an incremental no-op build is ~1s and a stale binary silently skews the comparison.
    tools["ours"] = ensure_built(OURS_BIN, ROOT, "--release", always=True)
    if not opts.skip_libreoffice:
        tools["soffice"] = find_soffice()
        if not tools["soffice"]:
            print("LibreOffice not found; skipping (brew install --cask libreoffice or LIBREOFFICE_PATH)")
    if not opts.skip_minipdf:
        tools["minipdf"] = find_minipdf()
        if not tools["minipdf"]:
            print("MiniPdf not found; skipping (cargo install minipdf-cli, or MINIPDF_BIN)")
    if not opts.skip_rdocx:
        tools["rdocx"] = find_rdocx()
        if not tools["rdocx"]:
            print("rdocx not found; skipping (cargo install rdocx, or RDOCX_BIN)")
    if not opts.skip_office2pdf:
        tools["office2pdf"] = find_office2pdf()
        if not tools["office2pdf"]:
            print("office2pdf not found; skipping (cargo install office2pdf-cli, or OFFICE2PDF_BIN)")
    if not opts.skip_jubarte:
        tools["jubarte"] = find_jubarte()
        if not tools["jubarte"]:
            print("jubarte not found; skipping (cargo install jubarte-redlines, or JUBARTE_BIN)")
    if not opts.skip_dxpdf:
        tools["dxpdf"] = find_dxpdf()
        if not tools["dxpdf"]:
            print("dxpdf not found; skipping (cargo install dxpdf, or DXPDF_BIN)")
    flat_font_dir()   # built once here, before the worker threads start converting
    if not opts.no_scores:
        # Release: SSIM over a 205-page fixture is painfully slow unoptimized. Always rebuilt (a
        # no-op when unchanged) so an edit to the scoring code is never scored with the old binary.
        if ensure_built(METRICS_BIN, ROOT / "tools", "--release", "--bin", "page-metrics", always=True):
            global SCORING_FP
            SCORING_FP = scoring_fingerprint()

    # Once, before the workers: each would otherwise run every engine's --version on its first case.
    for t in tools.values():
        if t:
            engine_id(t)
    fonts_fingerprint()

    groups = opts.group or GROUPS
    fixtures = [(g, d) for g in groups if (FIXTURES / g).is_dir()
                for d in sorted((FIXTURES / g).iterdir()) if d.is_dir()]
    if opts.case:
        fixtures = [(g, d) for g, d in fixtures if any(fnmatch(d.name, c) for c in opts.case)]
    if opts.shard:
        i, n = map(int, opts.shard.split("/"))
        fixtures = [(g, d) for g, d in fixtures if shard_of(g, d.name, n) == i]
    print(f"{len(fixtures)} fixtures, {opts.jobs} jobs" + (f", shard {opts.shard}" if opts.shard else ""))

    results: list[dict] = []
    with ThreadPoolExecutor(max_workers=opts.jobs) as pool:
        futures = {pool.submit(process_fixture, d, g, tools, opts): (g, d.name) for g, d in fixtures}
        # As they finish: in submission order one 200-page case stalls the log for every case behind it.
        for i, fut in enumerate(as_completed(futures), 1):
            g, name = futures[fut]
            try:
                r = fut.result()
            except Exception as e:  # keep going; one bad fixture should not kill the report
                print(f"  [{i}/{len(futures)}] {g}/{name}: ERROR {e}", file=sys.stderr)
                continue
            if r:
                redone = r.pop("redone", [])
                results.append(r)
                sc = " ".join(f"{k[:4]}=J{v.get('jaccard', 0):.0f}/S{v.get('ssim', 0):.0f}/T{v.get('text_boundary', 0):.0f}"
                              for k, v in r["scores"].items())
                again = "  | redone: " + ", ".join(f"{p} {e} ({w})" for p, e, w in redone) if redone else ""
                print(f"  [{i}/{len(futures)}] {g}/{name}  {sc}{again[:200]}")

    print()
    if not (opts.case or opts.group):   # a filtered run owns nothing but what it was asked for
        prune_work(fixtures)
    report_redone()
    report_timeouts()
    finish(results, engine_versions(tools))


if __name__ == "__main__":
    main()
