"""Compare numeric-suffix line breaks with the synthetic Word references.

Requires pypdfium2. Run with --engine <docxide-pdf> --out <directory>.
"""
import argparse
import json
from pathlib import Path
import subprocess
import pypdfium2 as pdfium


def lines(path):
    with pdfium.PdfDocument(path) as doc:
        page = doc[0]
        text = page.get_textpage()
        result = [line.strip() for line in text.get_text_range().splitlines() if line.strip()]
        text.close()
        page.close()
        return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--engine', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    args.out.mkdir(parents=True, exist_ok=True)
    root = Path(__file__).resolve().parents[1] / 'tests/fixtures/cases'
    rows = []
    for name, count in [('numeric-suffix-final-five', 1), ('numeric-suffix-final-four', 1), ('numeric-suffix-split-final', 1), ('numeric-suffix-budget-control', 2), ('numeric-suffix-short-final', 1)]:
        case = root / name
        output = args.out / (name + '.pdf')
        subprocess.run([str(args.engine.resolve()), str(case / 'input.docx'), str(output.resolve())],
                       check=True, capture_output=True)
        expected = lines(case / 'reference.pdf')
        actual = lines(output)
        assert len(expected) == count, (name, 'Word reference', expected)
        assert actual == expected, (name, actual, expected)
        rows.append({'case': name, 'lines': actual})
    (args.out / 'report.json').write_text(json.dumps(rows, indent=2), encoding='utf8')
    print('Five numeric-suffix Word controls passed')


if __name__ == '__main__':
    main()
