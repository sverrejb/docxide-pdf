#!/usr/bin/env python3
"""
Generate showcase images for the README.

Runs the test suite, picks all passing cases (SSIM >= threshold) sorted by
name, resizes their reference and generated page PNGs, saves them to
showcase/, and rewrites the <!-- showcase-start/end --> section in README.md.
"""
import csv
import subprocess
import sys
from pathlib import Path
from PIL import Image

ROOT = Path(__file__).parent.parent
SHOWCASE_DIR = ROOT / "showcase"
SSIM_CSV = ROOT / "tests/output/ssim_results.csv"
README = ROOT / "README.md"
TARGET_W = 420
SSIM_THRESHOLD = 0.40
IMG_BASE = "https://raw.githubusercontent.com/sverrejb/docxside-pdf/main/showcase"
START_MARKER = "<!-- showcase-start -->"
END_MARKER = "<!-- showcase-end -->"


def run_tests():
    print("Running tests...")
    result = subprocess.run(
        ["cargo", "test", "--", "--nocapture"],
        cwd=ROOT,
    )
    if result.returncode != 0:
        print("WARN: tests reported failures — using existing output", file=sys.stderr)


def passing_cases():
    if not SSIM_CSV.exists():
        print(f"No SSIM results at {SSIM_CSV}", file=sys.stderr)
        sys.exit(1)

    best = {}
    with open(SSIM_CSV) as f:
        for row in csv.DictReader(f):
            best[row["case"]] = float(row["avg_ssim"])

    passing = [(c, s) for c, s in best.items() if s >= SSIM_THRESHOLD]
    passing.sort(key=lambda x: x[0])
    return passing


def resize(src: Path, dst: Path):
    img = Image.open(src)
    ratio = TARGET_W / img.width
    img = img.resize((TARGET_W, int(img.height * ratio)), Image.LANCZOS)
    img.save(dst)


def build_section(rows):
    lines = ["<table>", "  <tr><th>MS Word</th><th>Docxside-PDF</th></tr>"]
    for case, score, ref_file, gen_file in rows:
        lines.append("  <tr>")
        lines.append(f'    <td align="center"><img src="{IMG_BASE}/{ref_file}"/><br/><sub>{case} — reference</sub></td>')
        lines.append(f'    <td align="center"><img src="{IMG_BASE}/{gen_file}"/><br/><sub>{case} — {score*100:.1f}% SSIM</sub></td>')
        lines.append("  </tr>")
    lines.append("</table>")
    return "\n".join(lines)


def update_readme(section):
    text = README.read_text()
    si = text.find(START_MARKER)
    ei = text.find(END_MARKER)
    if si == -1 or ei == -1:
        print("Showcase markers not found in README.md", file=sys.stderr)
        sys.exit(1)
    updated = f"{text[:si + len(START_MARKER)]}\n{section}\n{END_MARKER}{text[ei + len(END_MARKER):]}"
    README.write_text(updated)


def main():
    run_tests()

    cases = passing_cases()
    print(f"Passing cases (SSIM >= {SSIM_THRESHOLD*100:.0f}%):")
    for case, score in cases:
        print(f"  {case}: {score*100:.1f}%")

    SHOWCASE_DIR.mkdir(exist_ok=True)

    rows = []
    for case, score in cases:
        ref_src = ROOT / "tests/output" / case / "reference" / "page_001.png"
        gen_src = ROOT / "tests/output" / case / "generated" / "page_001.png"

        if not ref_src.exists() or not gen_src.exists():
            print(f"WARN: PNGs missing for {case}, skipping", file=sys.stderr)
            continue

        ref_dst = SHOWCASE_DIR / f"{case}_ref.png"
        gen_dst = SHOWCASE_DIR / f"{case}_gen.png"

        resize(ref_src, ref_dst)
        resize(gen_src, gen_dst)
        print(f"  Saved {ref_dst.name}")
        print(f"  Saved {gen_dst.name}")

        rows.append((case, score, ref_dst.name, gen_dst.name))

    update_readme(build_section(rows))
    print("README.md updated.")


if __name__ == "__main__":
    main()
