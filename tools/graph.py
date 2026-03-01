#!/usr/bin/env python3
import subprocess
from datetime import timedelta
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.animation as animation
import sys
from pathlib import Path

root = Path(__file__).parent.parent

jaccard_csv = root / "tests/output/results.csv"
ssim_csv = root / "tests/output/ssim_results.csv"

if not jaccard_csv.exists() and not ssim_csv.exists():
    print("No results yet — run cargo test first", file=sys.stderr)
    sys.exit(1)


def load_commits():
    log = subprocess.check_output(
        ["git", "log", "--pretty=format:%at\t%s"],
        cwd=root,
        text=True,
    )
    commits = []
    for line in log.splitlines():
        ts, _, msg = line.partition("\t")
        commits.append((pd.to_datetime(int(ts), unit="s"), msg))
    commits.sort(key=lambda c: c[0])
    return commits


def draw_commits(ax, commits):
    y_top = ax.get_ylim()[1]
    for i, (t, msg) in enumerate(commits):
        ax.axvline(t, color="gray", linewidth=0.6, alpha=0.35, linestyle="--")
        y_pos = y_top * (0.97 - 0.10 * (i % 4))
        ax.text(t, y_pos, msg, rotation=90, fontsize=6.5,
                va="top", ha="right", color="gray", alpha=0.55)


num_plots = int(jaccard_csv.exists()) + int(ssim_csv.exists())
fig, axes = plt.subplots(1, num_plots, figsize=(12 * num_plots, 6), squeeze=False)


def redraw(_frame=None):
    commits = load_commits()
    if not commits:
        return

    data_times = []
    for p, col in [(jaccard_csv, "avg_jaccard"), (ssim_csv, "avg_ssim")]:
        if p.exists():
            tmp = pd.read_csv(p)
            data_times.extend(pd.to_datetime(tmp["timestamp"], unit="s").tolist())

    t_first = min(data_times) if data_times else commits[0][0]
    t_last = commits[-1][0]
    padding = (t_last - t_first) * 0.03 or timedelta(minutes=5)

    plot_idx = 0

    if jaccard_csv.exists():
        df = pd.read_csv(jaccard_csv)
        df = df[df["case"].str.startswith("case")]
        df["time"] = pd.to_datetime(df["timestamp"], unit="s")
        ax = axes[0][plot_idx]
        ax.cla()
        for case, g in df.groupby("case"):
            ax.plot(g["time"], g["avg_jaccard"] * 100, marker="o", label=case)
        ax.axhline(25, linestyle="--", color="gray", linewidth=1, label="threshold (25%)")
        t_right = max(t_last + timedelta(hours=2), df["time"].max() + padding)
        ax.set_xlim(t_first - padding, t_right)
        draw_commits(ax, commits)
        ax.set_ylabel("Jaccard similarity (%)")
        ax.set_xlabel("Time")
        ax.set_title("Jaccard similarity over time")
        ax.legend()
        plot_idx += 1

    if ssim_csv.exists():
        df = pd.read_csv(ssim_csv)
        df = df[df["case"].str.startswith("case")]
        df["time"] = pd.to_datetime(df["timestamp"], unit="s")
        ax = axes[0][plot_idx]
        ax.cla()
        for case, g in df.groupby("case"):
            ax.plot(g["time"], g["avg_ssim"] * 100, marker="o", label=case)
        ax.axhline(40, linestyle="--", color="gray", linewidth=1, label="threshold (40%)")
        t_right = max(t_last + timedelta(hours=2), df["time"].max() + padding)
        ax.set_xlim(t_first - padding, t_right)
        draw_commits(ax, commits)
        ax.set_ylabel("SSIM (%)")
        ax.set_xlabel("Time")
        ax.set_title("SSIM over time")
        ax.legend()

    fig.autofmt_xdate()
    fig.tight_layout()


redraw()
ani = animation.FuncAnimation(fig, redraw, interval=3000, cache_frame_data=False)
plt.show()
