#!/usr/bin/env python3
"""Side-by-side engine comparison: Word reference | docxide-pdf | LibreOffice | MiniPdf | rdocx.

Reuses PNGs the test harness already produced under tests/output/<group>/<case>/
(reference/, generated/, libreoffice/) and only converts what is missing. Conversions and
screenshots are cached in comparison/work/. The viewer is a self-contained static site:
comparison/index.html plus lossless WebP page images, deployable as-is with
tools/deploy_comparison.sh (work/ is excluded by comparison/.gitignore).

Usage:
    python3 tools/engine_compare.py                 # every fixture with a reference.pdf
    python3 tools/engine_compare.py --case case41 --case 'case2*'   # exact or glob, repeatable
    python3 tools/engine_compare.py --group cases --open
    python3 tools/engine_compare.py --skip-libreoffice --no-scores
    python3 tools/engine_compare.py --html-only     # rebuild index.html from the cached manifest, no re-scoring

rdocx: `rdocx` on PATH (cargo install rdocx) or RDOCX_BIN.
MiniPdf: the Rust crate's CLI, `minipdf` on PATH (cargo install minipdf-cli) or MINIPDF_BIN.
The .NET engine is a different implementation and is deliberately not what we compare against.
"""
from __future__ import annotations

import argparse
import html
import json
import os
import re
import shutil
import subprocess
import sys
import webbrowser
from concurrent.futures import ThreadPoolExecutor
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
GROUPS = ["cases", "scraped", "new", "samples"]
DPI = "150"  # same as tests/common/mod.rs MUTOOL_DPI

# (key, label). Key doubles as the PNG directory name.
ENGINES = [
    ("reference", "Word"),
    ("generated", "docxide-pdf"),
    ("libreoffice", "LibreOffice"),
    ("minipdf", "MiniPdf (Rust)"),
    ("rdocx", "rdocx"),
]
COMPETITORS = [k for k, _ in ENGINES if k != "reference"]


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


def screenshot(pdf: Path, out_dir: Path) -> list[Path]:
    """Render every page to out_dir/page_NNN.png. Skips when PNGs are newer than the PDF."""
    existing = sorted(out_dir.glob("page_*.png"))
    if existing and all(p.stat().st_mtime >= pdf.stat().st_mtime for p in existing):
        return existing
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in existing:
        old.unlink()
    subprocess.run(
        ["mutool", "draw", "-F", "png", "-r", DPI, "-o", str(out_dir / "page_%03d.png"), str(pdf)],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False,
    )
    return sorted(out_dir.glob("page_*.png"))


def convert_ours(docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx):
        return True
    pdf.parent.mkdir(parents=True, exist_ok=True)
    r = subprocess.run([str(OURS_BIN), str(docx), str(pdf)], capture_output=True, text=True)
    return r.returncode == 0 and pdf.exists()


def convert_libreoffice(soffice: Path, docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx):
        return True
    out = pdf.parent
    out.mkdir(parents=True, exist_ok=True)
    # Per-case profile dir sidesteps LibreOffice's single-instance lock (same trick as the harness).
    profile = (out / "lo_profile").resolve()
    profile.mkdir(exist_ok=True)
    subprocess.run(
        [str(soffice), f"-env:UserInstallation=file://{profile}", "--headless",
         "--convert-to", "pdf", "--outdir", str(out), str(docx)],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300, check=False,
    )
    produced = out / (docx.stem + ".pdf")
    if produced.exists() and produced != pdf:
        produced.replace(pdf)
    return pdf.exists()


def convert_minipdf(minipdf: Path, docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx):
        return True
    pdf.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run([str(minipdf), "convert", str(docx), "-o", str(pdf)],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300, check=False)
    return pdf.exists()


def convert_rdocx(rdocx: Path, docx: Path, pdf: Path) -> bool:
    if is_fresh(pdf, docx):
        return True
    pdf.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run([str(rdocx), "convert", "--to", "pdf", "--output", str(pdf), str(docx)],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=300, check=False)
    return pdf.exists()


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
    for key in ("minipdf", "rdocx"):
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
    """Jaccard, SSIM and text-boundary in percent, computed by tools/page-metrics with the harness's own code."""
    if not METRICS_BIN.is_file():
        return {}
    r = subprocess.run([str(METRICS_BIN), str(ref_pdf), str(other_pdf), str(ref_dir), str(other_dir)],
                       capture_output=True, text=True)
    try:
        m = json.loads(r.stdout)
    except ValueError:
        return {}
    return {k: round(m[k] * 100, 1) for k in METRICS if m.get(k) is not None}


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

    def add(key: str, pdf: Path, png_dir: Path) -> None:
        pdfs[key] = pdf
        pages[key] = screenshot(pdf, png_dir)

    # Word reference
    ref_dir = harness / "reference" if any((harness / "reference").glob("page_*.png")) else mine / "reference"
    add("reference", ref_pdf, ref_dir)

    # Ours: prefer the harness output so scores line up with run-tests.sh
    if (harness / "generated.pdf").exists():
        add("generated", harness / "generated.pdf", harness / "generated")
    elif tools.get("ours") and convert_ours(docx, mine / "generated.pdf"):
        add("generated", mine / "generated.pdf", mine / "generated")

    if tools.get("soffice"):
        if (harness / "libreoffice.pdf").exists():
            add("libreoffice", harness / "libreoffice.pdf", harness / "libreoffice")
        elif convert_libreoffice(tools["soffice"], docx, mine / "libreoffice.pdf"):
            add("libreoffice", mine / "libreoffice.pdf", mine / "libreoffice")

    if tools.get("minipdf") and convert_minipdf(tools["minipdf"], docx, mine / "minipdf.pdf"):
        add("minipdf", mine / "minipdf.pdf", mine / "minipdf")

    if tools.get("rdocx") and convert_rdocx(tools["rdocx"], docx, mine / "rdocx.pdf"):
        add("rdocx", mine / "rdocx.pdf", mine / "rdocx")

    scores: dict[str, dict] = {}
    if not opts.no_scores:
        for key in COMPETITORS:
            if pages.get(key):
                m = engine_metrics(ref_pdf, pdfs[key], ref_dir, pages[key][0].parent)
                if m:
                    scores[key] = m

    rel = lambda p: os.path.relpath(p, WORK)  # noqa: E731
    return {
        "group": group,
        "case": case,
        "pages": {k: [rel(p) for p in v] for k, v in pages.items()},
        "scores": scores,
        "reference_app": pdf_creator(ref_pdf),
    }


HTML_TEMPLATE = r"""<!doctype html>
<meta charset="utf-8">
<title>Engine comparison</title>
<style>
:root { --bg:#1e1e1e; --panel:#252526; --fg:#ddd; --muted:#888; --accent:#4ea1ff; --border:#3a3a3a; }
* { box-sizing:border-box; }
body { margin:0; font:13px/1.4 -apple-system, Helvetica, Arial, sans-serif; background:var(--bg); color:var(--fg); display:grid; grid-template-columns:280px 1fr; grid-template-rows:auto auto 1fr; height:100vh; }
#legend { grid-column:1/3; padding:3px 10px 5px; font-size:11px; line-height:1.5; color:var(--muted); background:var(--panel); border-bottom:1px solid var(--border); }
#legend b { color:var(--fg); }
#legend span { margin-right:14px; }
#bar { grid-column:1/3; display:flex; gap:14px; align-items:center; padding:6px 10px; background:var(--panel); border-bottom:1px solid var(--border); flex-wrap:wrap; }
#bar label { display:inline-flex; align-items:center; gap:4px; cursor:pointer; user-select:none; }
#bar .grp { display:inline-flex; gap:8px; align-items:center; padding-right:14px; border-right:1px solid var(--border); }
#bar button, #bar select, #bar input[type=text] { background:#333; color:var(--fg); border:1px solid #555; border-radius:3px; padding:2px 8px; font:inherit; }
#bar button:hover { background:#444; }
kbd { background:#333; border:1px solid #555; border-radius:3px; padding:0 4px; font-size:11px; color:var(--muted); }
.ver { color:var(--muted); font-weight:400; font-size:11px; }
#side { overflow:auto; background:var(--panel); border-right:1px solid var(--border); }
#side input { width:100%; padding:6px 8px; background:#333; color:var(--fg); border:0; border-bottom:1px solid var(--border); font:inherit; position:sticky; top:0; }
#side .case { padding:4px 8px; cursor:pointer; border-bottom:1px solid #2e2e2e; display:grid; grid-template-columns:1fr auto; gap:0 8px; align-items:baseline; }
#side .case:hover { background:#2c2c2c; }
#side .case.sel { background:#094771; }
#side .name { font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
#side .grp { color:var(--muted); font-size:11px; }
#main { overflow:auto; padding:10px; }
#grid { display:grid; gap:4px 10px; align-items:start; }
.col { min-width:0; }
.colhead { margin:0; min-width:0; font-size:12px; font-weight:600; color:var(--muted); position:sticky; top:0; background:var(--bg); padding:2px 0; z-index:1; }
.colhead b { color:var(--fg); }
.page { margin-bottom:8px; }
.page img { width:100%; display:block; background:#fff; box-shadow:0 0 0 1px #000; }
.missing { width:100%; aspect-ratio:8.5/11; display:flex; align-items:center; justify-content:center; color:var(--muted); background:#2a2a2a; border:1px dashed #444; }
#overlay { position:relative; }
#overlay img { width:100%; display:block; }
#overlay img.top { position:absolute; inset:0; }
.hidden { display:none !important; }
#more { display:block; margin:4px 0 24px; padding:8px 18px; background:#333; color:var(--fg); border:1px solid #555; border-radius:4px; font:inherit; cursor:pointer; }
#more:hover { background:#444; }
</style>
<div id="bar">
  <span class="grp" id="engines"></span>
  <span class="grp">
    <span id="pageinfo"></span>
  </span>
  <span class="grp"><label>zoom <input type="range" id="zoom" min="200" max="1400" value="600" step="20"></label></span>
  <span class="grp">
    <label><input type="checkbox" id="ovl"> overlay</label>
    <select id="ovlA"></select> under <select id="ovlB"></select>
    <select id="blend"><option value="normal">opacity</option><option value="difference">difference</option><option value="multiply">multiply</option></select>
    <input type="range" id="alpha" min="0" max="100" value="50">
  </span>
  <span style="color:var(--muted)"><kbd>1</kbd>-<kbd>5</kbd> engines &nbsp;<kbd>&uarr;</kbd><kbd>&darr;</kbd> cases &nbsp;<kbd>o</kbd> overlay &nbsp;<kbd>m</kbd> more pages</span>
</div>
<div id="legend"></div>
<div id="side"><input id="filter" placeholder="filter cases (name, group)…"><div id="list"></div></div>
<div id="main"><div id="grid"></div><div id="overlay" class="hidden"></div><button id="more" class="hidden">more pages</button></div>
<script>
const DATA = __DATA__;
const ENGINES = __ENGINES__;
const METRICS = __METRICS__;
const VERSIONS = __VERSIONS__;
const METRIC_LABEL = { jaccard: 'J', ssim: 'SSIM', text_boundary: 'TB' };
const METRIC_INFO = {
  jaccard: 'Jaccard on ink pixels: both pages rendered at 150 DPI, a pixel is ink when its luma is below 200, score = ink in both ÷ ink in either. Exact placement matters: a one-line vertical shift sends it toward zero.',
  ssim: 'Structural similarity on 8×8 luma windows, each window allowed to search ±8 px vertically for its best match, so small vertical drift is forgiven. Only windows that contain ink count. Measures shape and texture rather than exact position.',
  text_boundary: 'Text boundary: share of text lines (mutool extraction) whose first and last word match the reference line at the same position. Pages whose line counts differ by more than 15% are skipped. Measures line breaking and pagination, independent of fonts and pixels.',
};
const PAGE_STEP = 10;
const $ = s => document.querySelector(s);
const store = k => { try { return JSON.parse(localStorage.getItem('ec.'+k)); } catch { return null; } };
const save = (k,v) => { try { localStorage.setItem('ec.'+k, JSON.stringify(v)); } catch {} };

let state = Object.assign({ on: {},
  shown: PAGE_STEP, zoom: 600, ovl: false, ovlA: 'reference', ovlB: 'generated', blend: 'normal', alpha: 50, sel: 0, filter: '' },
  store('state') || {});
// Engines added after a viewer state was saved default to visible.
for (const [k] of ENGINES) if (state.on[k] == null) state.on[k] = true;

// engine toggles
ENGINES.forEach(([key,label],i) => {
  const l = document.createElement('label');
  l.innerHTML = `<input type="checkbox" data-e="${key}"> ${label}${VERSIONS[key] ? ` <small class="ver">${VERSIONS[key]}</small>` : ''} <kbd>${i+1}</kbd>`;
  $('#engines').appendChild(l);
  ['#ovlA','#ovlB'].forEach(s => { const o = document.createElement('option'); o.value = key; o.textContent = label; $(s).appendChild(o); });
});
document.querySelectorAll('#engines input').forEach(cb => cb.onchange = () => { state.on[cb.dataset.e] = cb.checked; render(); });

function visibleCases() {
  const f = state.filter.toLowerCase();
  return DATA.map((c,i) => [c,i]).filter(([c]) => !f || (c.case + ' ' + c.group).toLowerCase().includes(f));
}
function fmt(v) { return v == null ? '–' : v.toFixed(1) + '%'; }

function renderList() {
  const list = $('#list'); list.innerHTML = '';
  for (const [c,i] of visibleCases()) {
    const d = document.createElement('div');
    d.className = 'case' + (i === state.sel ? ' sel' : '');
    d.innerHTML = `<span class="name" title="${c.case}">${c.case}</span><span class="grp">${c.group}</span>`;
    d.onclick = () => { state.sel = i; state.shown = PAGE_STEP; render(); };
    list.appendChild(d);
  }
}

function render() {
  save('state', state);
  document.querySelectorAll('#engines input').forEach(cb => cb.checked = !!state.on[cb.dataset.e]);
  $('#zoom').value = state.zoom; $('#ovl').checked = state.ovl;
  $('#ovlA').value = state.ovlA; $('#ovlB').value = state.ovlB; $('#blend').value = state.blend; $('#alpha').value = state.alpha;
  renderList();
  const c = DATA[state.sel]; if (!c) return;
  const maxPages = Math.max(1, ...Object.values(c.pages).map(p => p.length));
  state.shown = Math.min(Math.max(PAGE_STEP, state.shown), maxPages);
  const pageIdx = [...Array(state.shown).keys()];
  const left = maxPages - state.shown;
  $('#pageinfo').textContent = `pages 1–${state.shown} of ${maxPages}`;
  $('#more').classList.toggle('hidden', left <= 0);
  $('#more').textContent = `more pages (${left} left)`;

  const grid = $('#grid'), ov = $('#overlay');
  grid.classList.toggle('hidden', state.ovl); ov.classList.toggle('hidden', !state.ovl);
  if (state.ovl) {
    ov.style.width = state.zoom + 'px'; ov.innerHTML = '';
    for (const p of pageIdx) {
      const wrap = document.createElement('div'); wrap.style.position = 'relative'; wrap.style.marginBottom = '8px';
      const a = (c.pages[state.ovlA]||[])[p], b = (c.pages[state.ovlB]||[])[p];
      wrap.innerHTML = (a ? `<img src="${a}">` : '<div class="missing">no page</div>') +
        (b ? `<img class="top" src="${b}" style="opacity:${state.alpha/100};mix-blend-mode:${state.blend}">` : '');
      ov.appendChild(wrap);
    }
    return;
  }
  const on = ENGINES.filter(([k]) => state.on[k]);
  grid.style.gridTemplateColumns = `repeat(${on.length}, ${state.zoom}px)`;
  grid.innerHTML = '';
  // Headers occupy their own grid row so a header that wraps to two lines in one column
  // does not push that column's pages down relative to the others.
  const cols = [];
  for (const [key,label] of on) {
    const s = c.scores[key] || {}; const n = (c.pages[key]||[]).length;
    const sc = METRICS.filter(m => s[m] != null).map(m => ` · <span title="${METRIC_INFO[m]}">${METRIC_LABEL[m]} ${fmt(s[m])}</span>`).join('');
    // Reference PDFs printed via macOS (Producer "Quartz PDFContext") differ from Word's own export; flag them.
    const odd = key === 'reference' && c.reference_app && c.reference_app !== 'Microsoft Word' ? ` (${c.reference_app})` : '';
    const ver = (VERSIONS[key] || '') + odd;
    const h = document.createElement('h3'); h.className = 'colhead';
    h.innerHTML = `<b>${label}</b>${ver ? ` <span class="ver">${ver}</span>` : ''} · ${n} p${sc}`;
    grid.appendChild(h);
    const col = document.createElement('div'); col.className = 'col';
    for (const p of pageIdx) {
      const src = (c.pages[key]||[])[p];
      const d = document.createElement('div'); d.className = 'page';
      d.innerHTML = src ? `<img src="${src}" loading="lazy">` : `<div class="missing">no page ${p+1}</div>`;
      col.appendChild(d);
    }
    cols.push(col);
  }
  cols.forEach(col => grid.appendChild(col));
}

$('#more').onclick = () => { state.shown += PAGE_STEP; render(); };
$('#zoom').oninput = e => { state.zoom = +e.target.value; render(); };
$('#ovl').onchange = e => { state.ovl = e.target.checked; render(); };
$('#ovlA').onchange = e => { state.ovlA = e.target.value; render(); };
$('#ovlB').onchange = e => { state.ovlB = e.target.value; render(); };
$('#blend').onchange = e => { state.blend = e.target.value; render(); };
$('#alpha').oninput = e => { state.alpha = +e.target.value; render(); };
$('#legend').innerHTML = '<span>All scores are against the Word reference, averaged over the pages both PDFs have; extra pages are ignored.</span>' +
  METRICS.map(m => `<span><b>${METRIC_LABEL[m]}</b> ${METRIC_INFO[m]}</span>`).join('');
$('#filter').value = state.filter;
$('#filter').oninput = e => { state.filter = e.target.value; renderList(); };
document.onkeydown = e => {
  if (e.target.tagName === 'INPUT' && e.target.type === 'text') return;
  const vis = visibleCases().map(([,i]) => i); const at = vis.indexOf(state.sel);
  if (e.key === 'ArrowDown') { state.sel = vis[Math.min(at+1, vis.length-1)] ?? state.sel; state.shown = PAGE_STEP; }
  else if (e.key === 'ArrowUp') { state.sel = vis[Math.max(at-1, 0)] ?? state.sel; state.shown = PAGE_STEP; }
  else if (e.key === 'o') state.ovl = !state.ovl;
  else if (e.key === 'm') state.shown += PAGE_STEP;
  else if (/^[1-9]$/.test(e.key) && ENGINES[+e.key-1]) { const k = ENGINES[+e.key-1][0]; state.on[k] = !state.on[k]; }
  else return;
  e.preventDefault(); render();
  document.querySelector('#side .sel')?.scrollIntoView({block:'nearest'});
};
render();
</script>
"""


def natural_key(s: str) -> list:
    """case1 < case2 < case10, not case1 < case10 < case2."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


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
                src = (WORK / rel).resolve()
                dst = SITE / c["group"] / c["case"] / eng / (Path(rel).stem + "." + fmt)
                new.append(dst.relative_to(SITE).as_posix())
                if not dst.exists() or dst.stat().st_mtime < src.stat().st_mtime:
                    jobs_list.append((src, dst))
            pages[eng] = new
        rewritten.append({**c, "pages": pages})

    def transfer(pair: tuple[Path, Path]) -> None:
        src, dst = pair
        dst.parent.mkdir(parents=True, exist_ok=True)
        if fmt == "png":
            shutil.copy2(src, dst)
        else:
            subprocess.run(["cwebp", "-quiet", "-lossless", str(src), "-o", str(dst)],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)

    print(f"site: {len(jobs_list)} images to encode into {SITE}")
    with ThreadPoolExecutor(max_workers=jobs) as pool:
        list(pool.map(transfer, jobs_list))
    write_html(rewritten, versions, SITE / "index.html")
    (SITE / ".nojekyll").touch()  # GitHub Pages: serve as-is, no Jekyll pass over 9k files
    (SITE / ".gitignore").write_text("/work/\n")  # deploy_comparison.sh commits this folder; keep the cache out
    total = sum(f.stat().st_size for f in SITE.rglob("*") if f.is_file() and WORK not in f.parents)
    print(f"site ready: {SITE / 'index.html'} ({total / 1e6:.0f} MB)")


def write_html(results: list[dict], versions: dict, out: Path) -> None:
    results.sort(key=lambda r: (GROUPS.index(r["group"]), natural_key(r["case"])))
    out.parent.mkdir(parents=True, exist_ok=True)
    page = (HTML_TEMPLATE
            .replace("__DATA__", json.dumps(results))
            .replace("__ENGINES__", json.dumps(ENGINES))
            .replace("__METRICS__", json.dumps(METRICS))
            .replace("__VERSIONS__", json.dumps(versions)))
    out.write_text(page)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--case", action="append", default=[], help="case name or glob, e.g. case41 or 'case4*' (repeatable)")
    ap.add_argument("--group", choices=GROUPS, action="append", default=[], help="fixture group (repeatable)")
    ap.add_argument("--skip-libreoffice", action="store_true")
    ap.add_argument("--skip-minipdf", action="store_true")
    ap.add_argument("--skip-rdocx", action="store_true")
    ap.add_argument("--no-scores", action="store_true", help="skip Jaccard scoring")
    ap.add_argument("--jobs", type=int, default=os.cpu_count() or 4)
    ap.add_argument("--open", action="store_true", help="open the HTML when done")
    ap.add_argument("--html-only", action="store_true",
                    help="rebuild comparison/index.html from comparison/work/manifest.json (no conversion or scoring)")
    ap.add_argument("--format", choices=["webp", "png"], default="webp",
                    help="site image format; webp is lossless and ~3.5x smaller than png (needs cwebp)")
    opts = ap.parse_args()

    def load_manifest() -> dict:
        m = json.loads(MANIFEST.read_text()) if MANIFEST.exists() else {}
        return {"versions": {}, "cases": m} if isinstance(m, list) else m  # pre-versions manifests were a bare list

    def finish(results: list[dict], versions: dict) -> None:
        # A filtered run (--case/--group) updates just those entries; the site keeps every other case.
        if opts.case or opts.group:
            old = load_manifest()
            done = {(c["group"], c["case"]) for c in results}
            results = [c for c in old["cases"] if (c["group"], c["case"]) not in done] + results
            versions = {**old["versions"], **versions}
        results.sort(key=lambda r: (GROUPS.index(r["group"]), natural_key(r["case"])))
        WORK.mkdir(parents=True, exist_ok=True)
        MANIFEST.write_text(json.dumps({"versions": versions, "cases": results}))  # --html-only rebuilds from this
        print(f"{len(results)} cases; engines: " + ", ".join(f"{k} {v}" for k, v in versions.items()))
        build_site(results, versions, opts.format, opts.jobs)
        if opts.open:
            webbrowser.open((SITE / "index.html").as_uri())

    if opts.html_only:
        m = load_manifest()
        finish(m["cases"], m["versions"])
        return

    if not shutil.which("mutool"):
        sys.exit("mutool not found (brew install mupdf-tools)")
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
    if not opts.no_scores:
        # Release: SSIM over a 205-page fixture is painfully slow unoptimized.
        ensure_built(METRICS_BIN, ROOT / "tools", "--release", "--bin", "page-metrics")

    groups = opts.group or GROUPS
    fixtures = [(g, d) for g in groups if (FIXTURES / g).is_dir()
                for d in sorted((FIXTURES / g).iterdir()) if d.is_dir()]
    if opts.case:
        fixtures = [(g, d) for g, d in fixtures if any(fnmatch(d.name, c) for c in opts.case)]
    print(f"{len(fixtures)} fixtures, {opts.jobs} jobs")

    results: list[dict] = []
    with ThreadPoolExecutor(max_workers=opts.jobs) as pool:
        futures = {pool.submit(process_fixture, d, g, tools, opts): (g, d.name) for g, d in fixtures}
        for i, fut in enumerate(futures, 1):
            g, name = futures[fut]
            try:
                r = fut.result()
            except Exception as e:  # keep going; one bad fixture should not kill the report
                print(f"  [{i}/{len(futures)}] {g}/{name}: ERROR {e}", file=sys.stderr)
                continue
            if r:
                results.append(r)
                sc = " ".join(f"{k[:4]}=J{v.get('jaccard', 0):.0f}/S{v.get('ssim', 0):.0f}/T{v.get('text_boundary', 0):.0f}"
                              for k, v in r["scores"].items())
                print(f"  [{i}/{len(futures)}] {g}/{name}  {sc}")

    print()
    finish(results, engine_versions(tools))


if __name__ == "__main__":
    main()
