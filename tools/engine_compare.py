#!/usr/bin/env python3
"""Side-by-side engine comparison: Word reference | docxside-pdf | LibreOffice | MiniPdf.

Reuses PNGs the test harness already produced under tests/output/<group>/<case>/
(reference/, generated/, libreoffice/) and only converts what is missing. MiniPdf
output always lands under competitor/<group>/<case>/. Writes a single interactive
HTML viewer to competitor/compare.html.

Usage:
    python3 tools/engine_compare.py                 # every fixture with a reference.pdf
    python3 tools/engine_compare.py --case case41 --case 'case2*'   # exact or glob, repeatable
    python3 tools/engine_compare.py --group cases --open
    python3 tools/engine_compare.py --skip-libreoffice --no-scores

MiniPdf binary: MINIPDF_BIN env, competitor/minipdf/minipdf, or `minipdf` on PATH.
Install with:
    gh release download -R mini-software/MiniPdf -p 'minipdf-osx-arm64.tar.gz' -D competitor/minipdf
    tar xzf competitor/minipdf/minipdf-osx-arm64.tar.gz -C competitor/minipdf
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
COMPETITOR = ROOT / "competitor"
OURS_BIN = ROOT / "target" / "release" / "docxide-pdf"
JACCARD_BIN = ROOT / "tools" / "target" / "debug" / "jaccard"
GROUPS = ["cases", "scraped", "new", "samples"]
DPI = "150"  # same as tests/common/mod.rs MUTOOL_DPI

# (key, label). Key doubles as the PNG directory name.
ENGINES = [
    ("reference", "Word"),
    ("generated", "docxside-pdf"),
    ("libreoffice", "LibreOffice"),
    ("minipdf", "MiniPdf"),
]


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
    env = os.environ.get("MINIPDF_BIN")
    if env and Path(env).is_file():
        return Path(env)
    local = COMPETITOR / "minipdf" / "minipdf"
    if local.is_file():
        return local
    found = shutil.which("minipdf")
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


def jaccard_mean(ref_dir: Path, other_dir: Path) -> float | None:
    """Mean per-page Jaccard via tools/jaccard (same ink metric as the harness). Missing pages score 0."""
    if not JACCARD_BIN.is_file():
        return None
    r = subprocess.run([str(JACCARD_BIN), str(ref_dir), str(other_dir)], capture_output=True, text=True)
    scores = [float(m.group(1)) for m in re.finditer(r"^page_\d+\.png\s+([\d.]+)%", r.stdout, re.M)]
    n_ref = len(list(ref_dir.glob("page_*.png")))
    if not scores or n_ref == 0:
        return None
    return sum(scores) / max(n_ref, len(scores))


def process_fixture(fixture: Path, group: str, tools: dict, opts) -> dict | None:
    docx = fixture / "input.docx"
    ref_pdf = fixture / "reference.pdf"
    if not (docx.exists() and ref_pdf.exists()):
        return None
    case = fixture.name
    harness = TEST_OUTPUT / group / case
    mine = COMPETITOR / group / case
    pages: dict[str, list[Path]] = {}

    # Word reference
    ref_dir = harness / "reference" if any((harness / "reference").glob("page_*.png")) else mine / "reference"
    pages["reference"] = screenshot(ref_pdf, ref_dir)

    # Ours: prefer the harness output so scores line up with run-tests.sh
    ours_pdf = harness / "generated.pdf"
    if ours_pdf.exists():
        pages["generated"] = screenshot(ours_pdf, harness / "generated")
    elif tools.get("ours") and convert_ours(docx, mine / "generated.pdf"):
        pages["generated"] = screenshot(mine / "generated.pdf", mine / "generated")

    if tools.get("soffice"):
        lo_pdf = harness / "libreoffice.pdf"
        if lo_pdf.exists():
            pages["libreoffice"] = screenshot(lo_pdf, harness / "libreoffice")
        elif convert_libreoffice(tools["soffice"], docx, mine / "libreoffice.pdf"):
            pages["libreoffice"] = screenshot(mine / "libreoffice.pdf", mine / "libreoffice")

    if tools.get("minipdf") and convert_minipdf(tools["minipdf"], docx, mine / "minipdf.pdf"):
        pages["minipdf"] = screenshot(mine / "minipdf.pdf", mine / "minipdf")

    scores = {}
    if not opts.no_scores:
        for key in ("generated", "libreoffice", "minipdf"):
            if key in pages and pages[key]:
                s = jaccard_mean(pages["reference"][0].parent, pages[key][0].parent)
                if s is not None:
                    scores[key] = round(s, 1)

    rel = lambda p: os.path.relpath(p, COMPETITOR)  # noqa: E731
    return {
        "group": group,
        "case": case,
        "pages": {k: [rel(p) for p in v] for k, v in pages.items()},
        "scores": scores,
    }


def load_harness_scores() -> dict:
    try:
        return json.loads((TEST_OUTPUT / "latest_scores.json").read_text())
    except (OSError, ValueError):
        return {}


HTML_TEMPLATE = r"""<!doctype html>
<meta charset="utf-8">
<title>Engine comparison</title>
<style>
:root { --bg:#1e1e1e; --panel:#252526; --fg:#ddd; --muted:#888; --accent:#4ea1ff; --border:#3a3a3a; }
* { box-sizing:border-box; }
body { margin:0; font:13px/1.4 -apple-system, Helvetica, Arial, sans-serif; background:var(--bg); color:var(--fg); display:grid; grid-template-columns:280px 1fr; grid-template-rows:auto 1fr; height:100vh; }
#bar { grid-column:1/3; display:flex; gap:14px; align-items:center; padding:6px 10px; background:var(--panel); border-bottom:1px solid var(--border); flex-wrap:wrap; }
#bar label { display:inline-flex; align-items:center; gap:4px; cursor:pointer; user-select:none; }
#bar .grp { display:inline-flex; gap:8px; align-items:center; padding-right:14px; border-right:1px solid var(--border); }
#bar button, #bar select, #bar input[type=text] { background:#333; color:var(--fg); border:1px solid #555; border-radius:3px; padding:2px 8px; font:inherit; }
#bar button:hover { background:#444; }
kbd { background:#333; border:1px solid #555; border-radius:3px; padding:0 4px; font-size:11px; color:var(--muted); }
#side { overflow:auto; background:var(--panel); border-right:1px solid var(--border); }
#side input { width:100%; padding:6px 8px; background:#333; color:var(--fg); border:0; border-bottom:1px solid var(--border); font:inherit; position:sticky; top:0; }
#side .case { padding:5px 8px; cursor:pointer; border-bottom:1px solid #2e2e2e; display:grid; grid-template-columns:1fr auto; gap:2px 8px; }
#side .case:hover { background:#2c2c2c; }
#side .case.sel { background:#094771; }
#side .name { font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
#side .grp { color:var(--muted); font-size:11px; }
#side .sc { font-size:11px; color:var(--muted); grid-column:1/3; display:flex; gap:8px; }
#side .sc b { color:var(--fg); font-weight:500; }
#main { overflow:auto; padding:10px; }
#grid { display:grid; gap:10px; align-items:start; }
.col { min-width:0; }
.col h3 { margin:0 0 4px; font-size:12px; font-weight:600; color:var(--muted); position:sticky; top:0; background:var(--bg); padding:2px 0; }
.col h3 b { color:var(--fg); }
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
  <span style="color:var(--muted)"><kbd>1</kbd>-<kbd>4</kbd> engines &nbsp;<kbd>&uarr;</kbd><kbd>&darr;</kbd> cases &nbsp;<kbd>o</kbd> overlay &nbsp;<kbd>m</kbd> more pages</span>
</div>
<div id="side"><input id="filter" placeholder="filter cases (name, group)…"><div id="list"></div></div>
<div id="main"><div id="grid"></div><div id="overlay" class="hidden"></div><button id="more" class="hidden">more pages</button></div>
<script>
const DATA = __DATA__;
const ENGINES = __ENGINES__;
const HARNESS = __HARNESS__;
const PAGE_STEP = 10;
const $ = s => document.querySelector(s);
const store = k => { try { return JSON.parse(localStorage.getItem('ec.'+k)); } catch { return null; } };
const save = (k,v) => { try { localStorage.setItem('ec.'+k, JSON.stringify(v)); } catch {} };

let state = Object.assign({ on: {reference:true, generated:true, libreoffice:true, minipdf:true},
  shown: PAGE_STEP, zoom: 600, ovl: false, ovlA: 'reference', ovlB: 'generated', blend: 'normal', alpha: 50, sel: 0, filter: '' },
  store('state') || {});

// engine toggles
ENGINES.forEach(([key,label],i) => {
  const l = document.createElement('label');
  l.innerHTML = `<input type="checkbox" data-e="${key}"> ${label} <kbd>${i+1}</kbd>`;
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
    const h = HARNESS[c.group + '/' + c.case] || {};
    d.innerHTML = `<span class="name" title="${c.case}">${c.case}</span><span class="grp">${c.group}</span>
      <span class="sc">J: ours <b>${fmt(c.scores.generated)}</b> LO <b>${fmt(c.scores.libreoffice)}</b> Mini <b>${fmt(c.scores.minipdf)}</b>${h.ssim != null ? ` &nbsp;SSIM <b>${(h.ssim*100).toFixed(1)}%</b>` : ''}</span>`;
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
  for (const [key,label] of on) {
    const col = document.createElement('div'); col.className = 'col';
    const s = c.scores[key]; const n = (c.pages[key]||[]).length;
    col.innerHTML = `<h3><b>${label}</b> · ${n} p${s != null ? ` · J ${fmt(s)}` : ''}</h3>`;
    for (const p of pageIdx) {
      const src = (c.pages[key]||[])[p];
      const d = document.createElement('div'); d.className = 'page';
      d.innerHTML = src ? `<img src="${src}" loading="lazy">` : `<div class="missing">no page ${p+1}</div>`;
      col.appendChild(d);
    }
    grid.appendChild(col);
  }
}

$('#more').onclick = () => { state.shown += PAGE_STEP; render(); };
$('#zoom').oninput = e => { state.zoom = +e.target.value; render(); };
$('#ovl').onchange = e => { state.ovl = e.target.checked; render(); };
$('#ovlA').onchange = e => { state.ovlA = e.target.value; render(); };
$('#ovlB').onchange = e => { state.ovlB = e.target.value; render(); };
$('#blend').onchange = e => { state.blend = e.target.value; render(); };
$('#alpha').oninput = e => { state.alpha = +e.target.value; render(); };
$('#filter').value = state.filter;
$('#filter').oninput = e => { state.filter = e.target.value; renderList(); };
document.onkeydown = e => {
  if (e.target.tagName === 'INPUT' && e.target.type === 'text') return;
  const vis = visibleCases().map(([,i]) => i); const at = vis.indexOf(state.sel);
  if (e.key === 'ArrowDown') { state.sel = vis[Math.min(at+1, vis.length-1)] ?? state.sel; state.shown = PAGE_STEP; }
  else if (e.key === 'ArrowUp') { state.sel = vis[Math.max(at-1, 0)] ?? state.sel; state.shown = PAGE_STEP; }
  else if (e.key === 'o') state.ovl = !state.ovl;
  else if (e.key === 'm') state.shown += PAGE_STEP;
  else if (/^[1-4]$/.test(e.key)) { const k = ENGINES[+e.key-1][0]; state.on[k] = !state.on[k]; }
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


def write_html(results: list[dict], out: Path) -> None:
    results.sort(key=lambda r: (GROUPS.index(r["group"]), natural_key(r["case"])))
    # Manifest lets --html-only rebuild the page after template edits without re-scoring.
    out.with_suffix(".json").write_text(json.dumps(results))
    harness = load_harness_scores()
    page = (HTML_TEMPLATE
            .replace("__DATA__", json.dumps(results))
            .replace("__ENGINES__", json.dumps(ENGINES))
            .replace("__HARNESS__", json.dumps(harness)))
    out.write_text(page)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--case", action="append", default=[], help="case name or glob, e.g. case41 or 'case4*' (repeatable)")
    ap.add_argument("--group", choices=GROUPS, action="append", default=[], help="fixture group (repeatable)")
    ap.add_argument("--skip-libreoffice", action="store_true")
    ap.add_argument("--skip-minipdf", action="store_true")
    ap.add_argument("--no-scores", action="store_true", help="skip Jaccard scoring")
    ap.add_argument("--jobs", type=int, default=os.cpu_count() or 4)
    ap.add_argument("--open", action="store_true", help="open the HTML when done")
    ap.add_argument("--html-only", action="store_true", help="rebuild compare.html from the last run's manifest")
    opts = ap.parse_args()

    out = COMPETITOR / "compare.html"
    if opts.html_only:
        write_html(json.loads(out.with_suffix(".json").read_text()), out)
        print(f"wrote {out}")
        if opts.open:
            webbrowser.open(out.as_uri())
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
            print("MiniPdf not found; skipping (see --help for install)")
    if not opts.no_scores:
        ensure_built(JACCARD_BIN, ROOT / "tools", "--bin", "jaccard")

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
                sc = " ".join(f"{k[:4]}={v:.1f}" for k, v in r["scores"].items())
                print(f"  [{i}/{len(futures)}] {g}/{name}  {sc}")

    COMPETITOR.mkdir(exist_ok=True)
    write_html(results, out)
    print(f"\nwrote {out} ({len(results)} cases)")
    if opts.open:
        webbrowser.open(out.as_uri())


if __name__ == "__main__":
    main()
