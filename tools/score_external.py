#!/usr/bin/env python3
"""Score the current build on every document of the external corpus and print
one row per document, worst Jaccard first:

    python3 tools/score_external.py [--text] [--a11y] [--pages] [--sort METRIC] [--asc|--desc]
                                    [--cached] [--jobs N]

Scores every <group>/docx/<name>.docx against <group>/pdf/<name>.pdf under
~/workspace/docxside/corpora/neurotic_docx_bench with corpus_score.py's scorer, in the
All Markup view (--markup) that Word's local export prints.
Jaccard and SSIM only by default; --text adds text boundary and break drift,
--a11y the veraPDF/structure scores, --pages the page counts. Sorting on a
metric turns its group on. Results go to tests/output/score_external.json
(untracked); --cached re-sorts that file without converting anything.
Progress goes to stderr every 100 documents. The status column says why a
row has no scores: CONVERSION FAILED (no PDF; exit code or timeout), SCORER
FAILED, or NOT SCORED (usually our page size differs from Word's).
"""
from __future__ import annotations

import argparse
import itertools
import json
import os
import re
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from corpus_score import OUT, ROOT, score

CORPUS = ROOT.parent / 'corpora/neurotic_docx_bench'
RESULTS = ROOT / 'tests/output/score_external.json'
CLI = ROOT / 'target/release/docxide-pdf'
GROUPS = {
    'visual': ['jaccard', 'ssim'],
    'text': ['text_boundary', 'max_break_drift'],
    'a11y': ['ua_fail', 'ua_deficit', 'a11y_struct', 'a11y_text'],
    'pages': ['pages', 'ref_pages'],
}


def page_size(pdf: Path) -> str:
    out = subprocess.run(['pdfinfo', str(pdf)], capture_output=True, text=True).stdout
    m = re.search(r'Page size:\s+(\S+ x \S+)', out)
    return m.group(1) if m else '?'


def status(name: str, r: dict) -> str:
    if r.get('failed'):
        error = r.get('error', 'no PDF written')
        return f"{'SCORER' if error.startswith('page-metrics') else 'CONVERSION'} FAILED: {error}"
    if r.get('jaccard') is None:
        # The raster metrics only compare same-sized pages, so a page-size miss leaves nothing to score.
        group, stem = name.split('/', 1)
        ours = page_size(OUT / 'score_external' / group / stem / 'generated.pdf')
        word = page_size(CORPUS / group / 'pdf' / f'{stem}.pdf')
        return f'NOT SCORED: page size ours {ours} pt, Word {word} pt' if ours != word else 'NOT SCORED'
    return 'ok'


def build() -> None:
    for cmd, cwd in ((['cargo', 'build', '--release', '--bin', 'docxide-pdf'], ROOT),
                     (['cargo', 'build', '--release', '--bin', 'page-metrics'], ROOT / 'tools')):
        r = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
        if r.returncode:
            sys.exit(r.stderr)


def run(jobs: int, flags: tuple) -> dict:
    build()
    pairs = [(g.name, d, g / 'pdf' / f'{d.stem}.pdf')
             for g in sorted(CORPUS.iterdir()) if (g / 'docx').is_dir()
             for d in sorted((g / 'docx').glob('*.docx'))]
    pairs = [p for p in pairs if p[2].exists()]
    done = itertools.count(1)

    def one(p):
        # Every reference is Word for Mac's local export, which prints All Markup (deleted text,
        # revision balloons, comments); the final view loses pages and section breaks.
        s = score(p[1], p[2], str(CLI), OUT / 'score_external' / p[0], flags, ('--markup',))
        n = next(done)
        if n % 100 == 0 or n == len(pairs):
            print(f'{n}/{len(pairs)}', file=sys.stderr, flush=True)
        return s

    with ThreadPoolExecutor(jobs) as ex:
        results = {f'{g}/{d.stem}': s or {'failed': True} for (g, d, _), s in zip(pairs, ex.map(one, pairs))}
    RESULTS.write_text(json.dumps(results, indent=1))
    return results


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument('--text', action='store_true', help='also text boundary and break drift')
    p.add_argument('--a11y', action='store_true', help='also veraPDF and structure scores (slow)')
    p.add_argument('--pages', action='store_true', help='also page counts')
    p.add_argument('--sort', choices=[m for ms in GROUPS.values() for m in ms], default='jaccard')
    order = p.add_mutually_exclusive_group()
    order.add_argument('--asc', dest='desc', action='store_false', help='ascending (default)')
    order.add_argument('--desc', action='store_true')
    p.set_defaults(desc=False)
    p.add_argument('--cached', action='store_true', help='re-sort the last saved results')
    p.add_argument('--jobs', type=int, default=os.cpu_count())
    args = p.parse_args()
    on = {'visual', *(g for g in GROUPS if getattr(args, g, False)),
          next(g for g, ms in GROUPS.items() if args.sort in ms)}
    flags = tuple(f'--no-{g}' for g in ('text', 'a11y') if g not in on)
    results = json.loads(RESULTS.read_text()) if args.cached else run(args.jobs, flags)

    rows = [(name, {**r, **(r.get('a11y') or {})}) for name, r in results.items()]
    # A missing value (not scored, a11y not measured) ranks as the lowest, so worst-first lists it on top.
    rows.sort(key=lambda row: -float('inf') if row[1].get(args.sort) is None else row[1][args.sort],
              reverse=args.desc)
    cols = [m for g, ms in GROUPS.items() if g in on for m in ms]
    print('\t'.join(['document', *cols, 'status']))
    for name, r in rows:
        cells = [('-' if r.get(c) is None else f'{r[c] * 100:.2f}' if isinstance(r[c], float) else str(r[c]))
                 for c in cols]
        print('\t'.join([name, *cells, status(name, r)]))


if __name__ == '__main__':
    main()
