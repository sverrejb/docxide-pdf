#!/usr/bin/env python3
"""Compare baselines.json against latest_scores.json and print a compact diff.

Used by run-tests.sh to produce minimal output for LLM context windows.
No ANSI color codes — plain text only.
"""

import json
import sys
from pathlib import Path

REGRESSION_SLACK = 0.02
NOISE = 0.003
METRIC_NAMES = {"jaccard": "Jaccard", "ssim": "SSIM", "text_boundary": "TxtBnd"}
METRICS = ["jaccard", "ssim", "text_boundary"]


def short_name(name: str, max_len: int = 30) -> str:
    if len(name) > max_len:
        name = name[: max_len - 2] + ".."
    return name


def fmt_pct(v: float) -> str:
    return f"{v * 100:.1f}%"


def main():
    baselines_path = Path("tests/baselines.json")
    latest_path = Path("tests/output/latest_scores.json")

    if not latest_path.exists():
        print("No scores produced (compilation failure or no fixtures matched).")
        sys.exit(2)

    baselines = json.loads(baselines_path.read_text()) if baselines_path.exists() else {}
    latest = json.loads(latest_path.read_text())

    if not latest:
        print("No fixtures scored.")
        sys.exit(0)

    regressions = []
    improvements = []
    new_fixtures = []
    scored = 0

    for name, scores in sorted(latest.items()):
        scored += 1
        if name not in baselines:
            new_fixtures.append((name, scores))
            continue

        base = baselines[name]
        for m in METRICS:
            ov = base.get(m)
            nv = scores.get(m)
            if ov is None or nv is None:
                continue
            delta = nv - ov
            if abs(delta) < NOISE:
                continue
            entry = (name, m, ov, nv, delta)
            if delta < -REGRESSION_SLACK:
                regressions.append(entry)
            elif delta > REGRESSION_SLACK:
                improvements.append(entry)

    name_w = 30
    if regressions:
        regressions.sort(key=lambda x: x[4])
        print(f"Regressions ({len(regressions)}):")
        for name, m, ov, nv, delta in regressions:
            pp = delta * 100
            print(
                f"  {short_name(name):<{name_w}}  {METRIC_NAMES.get(m, m):<7}  "
                f"{fmt_pct(ov)} -> {fmt_pct(nv)}  {pp:+.1f}pp"
            )
        print()

    if improvements:
        improvements.sort(key=lambda x: x[4], reverse=True)
        print(f"Improvements ({len(improvements)}):")
        for name, m, ov, nv, delta in improvements:
            pp = delta * 100
            print(
                f"  {short_name(name):<{name_w}}  {METRIC_NAMES.get(m, m):<7}  "
                f"{fmt_pct(ov)} -> {fmt_pct(nv)}  {pp:+.1f}pp"
            )
        print()

    if new_fixtures:
        print(f"New fixtures ({len(new_fixtures)}):")
        for name, scores in new_fixtures:
            parts = []
            for m in METRICS:
                v = scores.get(m)
                if v is not None:
                    parts.append(f"{METRIC_NAMES.get(m, m)} {fmt_pct(v)}")
            print(f"  {short_name(name):<{name_w}}  {', '.join(parts)}")
        print()

    # Visual hash change detection
    visual_hashes_path = Path("tests/visual_hashes.json")
    latest_hashes_path = Path("tests/output/latest_hashes.json")
    hash_changed = []
    committed_h = {}
    if latest_hashes_path.exists():
        latest_h = json.loads(latest_hashes_path.read_text())
        if visual_hashes_path.exists():
            committed_h = json.loads(visual_hashes_path.read_text())
        for name, hashes in sorted(latest_h.items()):
            if name not in committed_h or committed_h[name] != hashes:
                hash_changed.append(name)

    if hash_changed:
        print(f"Visual changes ({len(hash_changed)}):")
        for name in hash_changed:
            status = "new" if name not in committed_h else "changed"
            print(f"  {short_name(name):<{name_w}}  {status}")
        print()

    changed = len(set(e[0] for e in regressions + improvements)) + len(new_fixtures)
    reg_count = len(set(e[0] for e in regressions))
    imp_count = len(set(e[0] for e in improvements))
    unchanged = scored - changed

    summary = f"{scored} scored, {unchanged} unchanged"
    if reg_count:
        summary += f", {reg_count} regressed"
    if imp_count:
        summary += f", {imp_count} improved"
    if new_fixtures:
        summary += f", {len(new_fixtures)} new"
    if hash_changed:
        summary += f", {len(hash_changed)} visual changes"
    print(summary)

    sys.exit(1 if regressions else 0)


if __name__ == "__main__":
    main()
