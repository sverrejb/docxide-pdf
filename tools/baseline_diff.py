#!/usr/bin/env python3
"""Compare current baselines.json against a previous git commit.

Usage:
    python3 tools/baseline_diff.py <commit>
    python3 tools/baseline_diff.py 93db3cd
    python3 tools/baseline_diff.py HEAD~5
"""

import json
import subprocess
import sys

THRESHOLDS = {"jaccard": 0.20, "ssim": 0.75}
METRIC_NAMES = {"jaccard": "Jaccard", "ssim": "SSIM", "text_boundary": "TxtBnd"}

# Score changes smaller than this are ignored
NOISE = 0.003

GREEN = "\033[32m"
RED = "\033[31m"
YELLOW = "\033[33m"
DIM = "\033[2m"
BOLD = "\033[1m"
RESET = "\033[0m"


def load_baseline_at(commit: str) -> dict:
    result = subprocess.run(
        ["git", "show", f"{commit}:tests/baselines.json"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(f"Error: could not read baselines.json at {commit}", file=sys.stderr)
        print(result.stderr.strip(), file=sys.stderr)
        sys.exit(1)
    return json.loads(result.stdout)


def load_current_baseline() -> dict:
    with open("tests/baselines.json") as f:
        return json.load(f)


def fmt_pct(v: float) -> str:
    return f"{v * 100:.1f}%"


def fmt_delta(d: float) -> str:
    pct = d * 100
    if pct > 0:
        return f"{GREEN}+{pct:.1f}%{RESET}"
    elif pct < 0:
        return f"{RED}{pct:.1f}%{RESET}"
    return f"{DIM}  0.0%{RESET}"


def short_name(name: str, max_len: int = 38) -> str:
    name = name.replace("scraped/", "").replace("cases/", "").replace("samples/", "s/")
    if len(name) > max_len:
        name = name[:max_len - 2] + ".."
    return name


def main():
    if len(sys.argv) < 2:
        print(__doc__.strip())
        sys.exit(1)

    commit = sys.argv[1]

    commit_label = subprocess.run(
        ["git", "log", "--oneline", "-1", commit],
        capture_output=True, text=True,
    ).stdout.strip()

    old = load_baseline_at(commit)
    new = load_current_baseline()

    all_keys = sorted(set(old) | set(new))
    metrics = ["jaccard", "ssim", "text_boundary"]

    improvements = []
    regressions = []
    added = []
    removed = []

    for key in all_keys:
        if key not in old:
            added.append(key)
            continue
        if key not in new:
            removed.append(key)
            continue

        for m in metrics:
            ov = old[key].get(m)
            nv = new[key].get(m)
            if ov is None or nv is None:
                continue
            delta = nv - ov
            if abs(delta) < NOISE:
                continue
            entry = (key, m, ov, nv, delta)
            if delta > 0:
                improvements.append(entry)
            else:
                regressions.append(entry)

    # Print header
    print(f"\n{BOLD}Baseline diff: {commit_label} → HEAD{RESET}\n")

    name_w = 38
    metric_w = 7
    val_w = 7
    delta_w = 16  # includes ANSI codes

    def print_header():
        print(f"  {'Fixture':<{name_w}}  {'Metric':<{metric_w}}  {'Old':>{val_w}}  {'New':>{val_w}}  {'Delta':>{8}}")
        print(f"  {'─' * name_w}  {'─' * metric_w}  {'─' * val_w}  {'─' * val_w}  {'─' * 8}")

    # Improvements
    if improvements:
        improvements.sort(key=lambda x: x[4], reverse=True)
        print(f"{GREEN}{BOLD}Improvements ({len(improvements)}){RESET}")
        print_header()
        for key, m, ov, nv, delta in improvements:
            print(f"  {short_name(key):<{name_w}}  {METRIC_NAMES.get(m, m):<{metric_w}}  {fmt_pct(ov):>{val_w}}  {fmt_pct(nv):>{val_w}}  {fmt_delta(delta)}")
        print()

    # Regressions
    if regressions:
        regressions.sort(key=lambda x: x[4])
        print(f"{RED}{BOLD}Regressions ({len(regressions)}){RESET}")
        print_header()
        for key, m, ov, nv, delta in regressions:
            print(f"  {short_name(key):<{name_w}}  {METRIC_NAMES.get(m, m):<{metric_w}}  {fmt_pct(ov):>{val_w}}  {fmt_pct(nv):>{val_w}}  {fmt_delta(delta)}")
        print()

    # New fixtures
    if added:
        print(f"{YELLOW}{BOLD}New fixtures ({len(added)}){RESET}")
        for key in added:
            vals = new[key]
            j = fmt_pct(vals.get("jaccard", 0))
            s = fmt_pct(vals.get("ssim", 0))
            print(f"  {short_name(key):<{name_w}}  Jaccard {j:>6}  SSIM {s:>6}")
        print()

    # Removed fixtures
    if removed:
        print(f"{DIM}Removed fixtures ({len(removed)}){RESET}")
        for key in removed:
            print(f"  {short_name(key)}")
        print()

    # Summary
    if not improvements and not regressions and not added and not removed:
        print(f"{DIM}No changes.{RESET}\n")
    else:
        total_j_delta = sum(d for _, m, _, _, d in improvements + regressions if m == "jaccard")
        total_s_delta = sum(d for _, m, _, _, d in improvements + regressions if m == "ssim")
        print(f"{BOLD}Net change across existing fixtures:{RESET}  "
              f"Jaccard {fmt_delta(total_j_delta)}  SSIM {fmt_delta(total_s_delta)}")
        print()


if __name__ == "__main__":
    main()
