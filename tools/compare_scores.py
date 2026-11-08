#!/usr/bin/env python3
"""Compare two score files (latest_scores.json format): per-group mean Jaccard
and SSIM, then every case that moved by more than the threshold.

    python3 tools/compare_scores.py <before.json> <after.json> [threshold]

The threshold is a fraction (default 0.005 = half a point).
"""
import collections
import json
import statistics
import sys


def main() -> None:
    a, b = json.load(open(sys.argv[1])), json.load(open(sys.argv[2]))
    thr = float(sys.argv[3]) if len(sys.argv) > 3 else 0.005
    keys = sorted(k for k in set(a) & set(b) if a[k] and b[k])
    groups = collections.defaultdict(list)
    for k in keys:
        groups[k.split('/')[0]].append(k)

    def mean(d, ks, field):
        return statistics.mean((d[k].get(field) or 0) for k in ks) * 100

    for g, ks in sorted(groups.items()):
        print(f"{g:11s} n={len(ks):3d}  J {mean(a, ks, 'jaccard'):5.2f} -> {mean(b, ks, 'jaccard'):5.2f}   "
              f"SSIM {mean(a, ks, 'ssim'):5.2f} -> {mean(b, ks, 'ssim'):5.2f}")
    changes = []
    for k in keys:
        dj = (b[k].get('jaccard') or 0) - (a[k].get('jaccard') or 0)
        ds = (b[k].get('ssim') or 0) - (a[k].get('ssim') or 0)
        if abs(dj) > thr or abs(ds) > thr:
            changes.append((dj, ds, k))
    for dj, ds, k in sorted(changes):
        print(f"  {k:45s} J {(a[k].get('jaccard') or 0) * 100:5.1f} -> {(b[k].get('jaccard') or 0) * 100:5.1f} ({dj * 100:+5.1f})"
              f"  SSIM {(a[k].get('ssim') or 0) * 100:5.1f} -> {(b[k].get('ssim') or 0) * 100:5.1f} ({ds * 100:+5.1f})")
    print(f"missing in after: {sorted(set(a) - set(b))}")


if __name__ == '__main__':
    main()
