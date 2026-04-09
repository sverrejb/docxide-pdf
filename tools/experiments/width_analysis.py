#!/usr/bin/env python3
"""Analyze per-character advance width differences between our raw font metrics
and Word's effective advances (from reference PDF positions).

Uses mutool stext to extract character positions from reference PDFs,
computes implied advances, and compares against fontTools metrics.
"""
import subprocess, sys, xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

try:
    from fontTools.ttLib import TTFont
except ImportError:
    print("pip install fonttools first")
    sys.exit(1)

FONT_PATHS = {
    "Calibri": "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Calibri.ttf",
    "TimesNewRomanPSMT": "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Times New Roman.ttf",
    "Times New Roman": "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Times New Roman.ttf",
    "Arial": "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Arial.ttf",
    "ArialMT": "/Applications/Microsoft Word.app/Contents/Resources/DFonts/Arial.ttf",
}

def load_font_widths(font_path, font_size):
    """Return {char: advance_in_points} from raw font metrics."""
    font = TTFont(font_path)
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    upm = font["head"].unitsPerEm
    widths = {}
    for cp in range(32, 0x2000):
        gname = cmap.get(cp)
        if gname:
            raw = hmtx[gname][0]
            widths[chr(cp)] = raw / upm * font_size
    return widths

def extract_char_positions(pdf_path, page=1):
    """Extract per-char (x, char, font_name, font_size) from mutool stext."""
    result = subprocess.run(
        ["mutool", "draw", "-F", "stext", str(pdf_path), str(page)],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"mutool failed: {result.stderr}")
        return []

    try:
        root = ET.fromstring(result.stdout)
    except ET.ParseError:
        return []

    chars = []
    for page_el in root.findall(".//page"):
        for line_el in page_el.findall(".//line"):
            for font_el in line_el.findall("font"):
                fname = font_el.get("name", "")
                fsize = float(font_el.get("size", "12"))
                for char_el in font_el.findall("char"):
                    x = float(char_el.get("x"))
                    y = float(char_el.get("y"))
                    c = char_el.get("c")
                    chars.append((x, y, c, fname, fsize))
    return chars

def analyze_advances(chars, font_widths_cache):
    """Compare implied advances from PDF positions against raw font metrics."""
    results = defaultdict(list)  # (font, size) -> [(char, word_advance, our_advance)]

    for i in range(len(chars) - 1):
        x1, y1, c1, fname1, fsize1 = chars[i]
        x2, y2, c2, fname2, fsize2 = chars[i + 1]

        # Only compare consecutive chars on the same line (same y, same font)
        if abs(y1 - y2) > 0.5 or fname1 != fname2 or abs(fsize1 - fsize2) > 0.1:
            continue

        # Skip spaces (may have justification)
        if c1 == " ":
            continue

        word_advance = x2 - x1
        if word_advance <= 0 or word_advance > fsize1 * 2:
            continue  # skip anomalies

        key = (fname1, fsize1)
        if key not in font_widths_cache:
            fpath = None
            for name_prefix, path in FONT_PATHS.items():
                if fname1.startswith(name_prefix) or name_prefix.startswith(fname1):
                    fpath = path
                    break
            if not fpath or not Path(fpath).exists():
                continue
            font_widths_cache[key] = load_font_widths(fpath, fsize1)

        our_advance = font_widths_cache[key].get(c1, None)
        if our_advance is None:
            continue

        diff = our_advance - word_advance
        results[key].append((c1, word_advance, our_advance, diff))

    return results

def main():
    fixtures = [
        "tests/fixtures/cases/case4/reference.pdf",
        "tests/fixtures/cases/case54/reference.pdf",
    ]

    # Also check for russian essay
    russian = Path("tests/fixtures/scraped/russian_volunteerism_essay/reference.pdf")
    if russian.exists():
        fixtures.append(str(russian))

    font_widths_cache = {}

    for pdf_path in fixtures:
        if not Path(pdf_path).exists():
            print(f"Skipping {pdf_path} (not found)")
            continue

        print(f"\n{'='*70}")
        print(f"Analyzing: {pdf_path}")
        print(f"{'='*70}")

        all_results = defaultdict(list)
        for page in range(1, 4):  # first 3 pages
            chars = extract_char_positions(pdf_path, page)
            if not chars:
                continue
            results = analyze_advances(chars, font_widths_cache)
            for key, vals in results.items():
                all_results[key].extend(vals)

        for (fname, fsize), entries in sorted(all_results.items()):
            diffs = [d for (_, _, _, d) in entries]
            if not diffs:
                continue

            n = len(diffs)
            mean_diff = sum(diffs) / n
            abs_diffs = [abs(d) for d in diffs]
            mean_abs = sum(abs_diffs) / n
            positive = sum(1 for d in diffs if d > 0.001)
            negative = sum(1 for d in diffs if d < -0.001)
            near_zero = n - positive - negative

            # Compute correction factor
            avg_advance = sum(our for (_, _, our, _) in entries) / n
            if avg_advance > 0:
                factor = 1.0 - mean_diff / avg_advance

            print(f"\n  {fname} {fsize}pt: {n} character advances")
            print(f"    Mean signed diff:   {mean_diff:+.4f}pt (ours wider)")
            print(f"    Mean absolute diff: {mean_abs:.4f}pt")
            print(f"    Positive/negative/zero: {positive}/{negative}/{near_zero}")
            print(f"    Avg advance width:  {avg_advance:.3f}pt")
            print(f"    Implied factor:     {factor:.6f}")
            print(f"    Cum error over 80 chars: {mean_diff * 80:.2f}pt")

            # Per-character breakdown (top diffs)
            char_diffs = defaultdict(list)
            for (c, wa, oa, d) in entries:
                char_diffs[c].append(d)

            print(f"    Per-char averages (top 10 by magnitude):")
            char_avgs = [(c, sum(ds)/len(ds), len(ds)) for c, ds in char_diffs.items()]
            char_avgs.sort(key=lambda x: abs(x[1]), reverse=True)
            for c, avg, cnt in char_avgs[:10]:
                print(f"      '{c}': {avg:+.4f}pt  (n={cnt})")

if __name__ == "__main__":
    main()
