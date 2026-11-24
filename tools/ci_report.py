#!/usr/bin/env python3
"""Markdown summary of an engine comparison run, compared with the previous published run.

    python3 tools/ci_report.py <previous report.json> <manifest.json> <new report.json>

Prints the summary on stdout (CI appends it to $GITHUB_STEP_SUMMARY) and writes a slim
report.json, published with the site, for the next run to compare against. A missing or
unreadable previous report just leaves out the deltas.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

MOVE = 1.0   # Jaccard points a case must move to count as up/down


def slim(manifest: dict) -> dict:
    cases = {f'{c["group"]}/{c["case"]}': {"scores": {e: {"jaccard": s.get("jaccard"), "ssim": s.get("ssim")}
                                                      for e, s in c.get("scores", {}).items()},
                                           "times": c.get("times", {}),
                                           "engines": [e for e, p in c.get("pages", {}).items() if p]}
             for c in manifest["cases"]}
    return {"versions": manifest["versions"], "cases": cases}


def mean(xs: list) -> float | None:
    return sum(xs) / len(xs) if xs else None


def fmt(x, delta=False) -> str:
    if x is None:
        return "–"
    return f"{x:+.2f}" if delta else f"{x:.1f}"


def jac(report: dict, case: str, engine: str):
    return report["cases"].get(case, {}).get("scores", {}).get(engine, {}).get("jaccard")


def version_cell(old: str | None, new: str) -> str:
    if old is None or old == new:
        return new
    repo = os.environ.get("GITHUB_REPOSITORY")
    if repo and "@" in old and "@" in new:   # our own build: link the commit range
        a, b = (v.split("@")[1].removesuffix("+dirty") for v in (old, new))
        return f"[{old} → {new}](https://github.com/{repo}/compare/{a}...{b})"
    return f"**{old} → {new}**"


def report(prev: dict, cur: dict) -> str:
    out = []
    cases, old_cases = set(cur["cases"]), set(prev.get("cases", {}))
    line = f"{len(cases)} cases"
    if old_cases:
        added, removed = sorted(cases - old_cases), sorted(old_cases - cases)
        line += f" ({len(added)} new, {len(removed)} removed since the last run)"
        line += "".join(f"\n- {label}: {', '.join(xs[:20])}{' …' if len(xs) > 20 else ''}"
                        for label, xs in (("New", added), ("Removed", removed)) if xs)
    else:
        line += " (no previous report to compare with)"
    out.append(line)

    old_v = prev.get("versions", {})
    changed = [e for e in cur["versions"] if e in old_v and old_v[e] != cur["versions"][e] and e != "generated"]
    if changed:
        out.append("\n**Engine versions changed:** " + ", ".join(f"{e} {old_v[e]} → {cur['versions'][e]}" for e in changed))

    common = cases & old_cases
    out.append("\n| Engine | Version | Converted | Mean Jaccard | Δ (same cases) | ↑ / ↓ cases | Mean SSIM | Mean s/doc |")
    out.append("|---|---|---|---|---|---|---|---|")
    for e, v in cur["versions"].items():
        if e == "reference":
            out.append(f"| {e} | {v} | | | | | | |")
            continue
        conv = sum(e in c["engines"] for c in cur["cases"].values())
        j = mean([x for c in cases if (x := jac(cur, c, e)) is not None])
        ssim = mean([s for c in cur["cases"].values() if (s := c["scores"].get(e, {}).get("ssim")) is not None])
        t = mean([s for c in cur["cases"].values() if (s := c["times"].get(e)) is not None])
        pairs = [(a, b) for c in common if (a := jac(prev, c, e)) is not None and (b := jac(cur, c, e)) is not None]
        d = mean([b for _, b in pairs]) - mean([a for a, _ in pairs]) if pairs else None
        up = sum(b - a >= MOVE for a, b in pairs)
        down = sum(a - b >= MOVE for a, b in pairs)
        old_conv = sum(e in c.get("engines", []) for c in prev.get("cases", {}).values())
        conv_cell = f"{conv}/{len(cases)}" + (f" (was {old_conv}/{len(old_cases)})" if old_cases and old_conv != conv else "")
        moves = f"{up} / {down}" if pairs else "–"
        out.append(f"| {e} | {version_cell(old_v.get(e), v)} | {conv_cell} | {fmt(j)} | {fmt(d, True)} | {moves} "
                   f"| {fmt(ssim)} | {fmt(t)} |")

    movers = sorted(((b - a, c, a, b) for c in common
                     if (a := jac(prev, c, "generated")) is not None and (b := jac(cur, c, "generated")) is not None
                     and abs(b - a) >= MOVE), key=lambda m: -abs(m[0]))
    if movers:
        out.append(f"\n<details><summary>docxide-pdf: {len(movers)} cases moved ≥ {MOVE} Jaccard</summary>\n")
        out.append("| Case | Before | After | Δ |\n|---|---|---|---|")
        out += [f"| {c} | {a:.1f} | {b:.1f} | {d:+.1f} |" for d, c, a, b in movers[:30]]
        out.append("\n</details>")
    return "\n".join(out)


from engine_compare import ENGINES

NAMES = dict(ENGINES)


def scores_svg(cur: dict) -> str:
    """Horizontal bar chart of each engine's mean Jaccard and SSIM, embedded in the README."""
    cases = list(cur["cases"].values())
    rows = []
    for e in cur["versions"]:
        if e == "reference":
            continue
        # A case the engine failed to convert counts as 0, so skipping hard documents earns nothing.
        j, s = (sum(c["scores"].get(e, {}).get(m) or 0 for c in cases) / len(cases) for m in ("jaccard", "ssim"))
        conv = sum(e in c["engines"] for c in cases)
        rows.append((NAMES.get(e, e), j, s, conv))
    rows.sort(key=lambda r: -r[1])

    label_w, plot_w, row_h, bar_h, top = 190, 440, 44, 14, 56
    w, h = label_w + plot_w + 50, top + row_h * len(rows) + 28
    x = lambda v: label_w + plot_w * v / 100
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}" '
           'font-family="-apple-system,Segoe UI,Helvetica,Arial,sans-serif" font-size="13">',
           '<style>:root{--bg:#fcfcfb;--ink:#0b0b0b;--ink2:#52514e;--grid:#e1e0d9;--j:#2a78d6;--s:#eb6834}'
           '@media (prefers-color-scheme:dark){:root{--bg:#1a1a19;--ink:#f0efec;--ink2:#c3c2b7;--grid:#2c2c2a;'
           '--j:#3987e5;--s:#d95926}}text{fill:var(--ink)}.m{fill:var(--ink2)}</style>',
           f'<rect width="{w}" height="{h}" rx="6" fill="var(--bg)"/>',
           f'<text x="16" y="24" font-weight="600">Mean score vs the Word reference, {len(cases)} documents</text>',
           f'<rect x="16" y="36" width="10" height="10" rx="2" fill="var(--j)"/><text x="31" y="45" class="m">Jaccard</text>',
           f'<rect x="96" y="36" width="10" height="10" rx="2" fill="var(--s)"/><text x="111" y="45" class="m">SSIM</text>']
    for t in range(0, 101, 25):
        out.append(f'<line x1="{x(t)}" x2="{x(t)}" y1="{top - 4}" y2="{h - 24}" stroke="var(--grid)"/>'
                   f'<text x="{x(t)}" y="{h - 8}" text-anchor="middle" class="m" font-size="11">{t}</text>')
    for i, (name, j, s, conv) in enumerate(rows):
        y = top + i * row_h
        failed = len(cases) - conv
        out.append(f'<text x="{label_w - 12}" y="{y + 15 if failed else y + 22}" text-anchor="end">{name}</text>')
        if failed:
            out.append(f'<text x="{label_w - 12}" y="{y + 30}" text-anchor="end" class="m" font-size="11">'
                       f'{failed} failed conversion{"s" if failed > 1 else ""}</text>')
        for k, (v, color, metric) in enumerate(((j, "j", "Jaccard"), (s, "s", "SSIM"))):
            by = y + 4 + k * (bar_h + 2)
            out.append(f'<rect x="{label_w}" y="{by}" width="{max(x(v) - label_w, 1):.1f}" height="{bar_h}" rx="3" '
                       f'fill="var(--{color})"><title>{name} {metric}: {v:.1f}</title></rect>'
                       f'<text x="{x(v) + 6:.1f}" y="{by + 11}" class="m" font-size="11">{v:.1f}</text>')
    out.append("</svg>")
    return "\n".join(out)


def main() -> None:
    prev_path, manifest_path, out_path = map(Path, sys.argv[1:4])
    try:
        prev = json.loads(prev_path.read_text())
    except (OSError, ValueError):
        prev = {}
    cur = slim(json.loads(manifest_path.read_text()))
    out_path.write_text(json.dumps(cur, separators=(",", ":")))
    out_path.with_name("scores.svg").write_text(scores_svg(cur))
    print("## Engine comparison\n")
    print(report(prev, cur))


if __name__ == "__main__":
    main()
