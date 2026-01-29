#!/usr/bin/env python3
import subprocess
from datetime import timedelta  # used for padding calculation
import pandas as pd
import matplotlib.pyplot as plt
import sys
from pathlib import Path

csv = Path(__file__).parent.parent / "tests/output/results.csv"
if not csv.exists():
    print(f"No results yet: {csv}", file=sys.stderr)
    sys.exit(1)

df = pd.read_csv(csv)
df["time"] = pd.to_datetime(df["timestamp"], unit="s")

# Commit timeline is the primary axis
log = subprocess.check_output(
    ["git", "log", "--pretty=format:%at\t%s"],
    cwd=Path(__file__).parent.parent,
    text=True,
)
commits = []
for line in log.splitlines():
    ts, _, msg = line.partition("\t")
    commits.append((pd.to_datetime(int(ts), unit="s"), msg))
commits.sort(key=lambda c: c[0])  # oldest first

if not commits:
    print("No commits found", file=sys.stderr)
    sys.exit(1)

t_first, t_last = commits[0][0], commits[-1][0]
padding = (t_last - t_first) * 0.03 or timedelta(minutes=5)

fig, ax = plt.subplots(figsize=(12, 6))

for case, g in df.groupby("case"):
    ax.plot(g["time"], g["avg_jaccard"] * 100, marker="o", label=case)

ax.axhline(40, linestyle="--", color="gray", linewidth=1, label="threshold (40%)")

y_top = ax.get_ylim()[1]
for i, (t, msg) in enumerate(commits):
    ax.axvline(t, color="gray", linewidth=0.6, alpha=0.35, linestyle="--")
    y_pos = y_top * (0.97 - 0.10 * (i % 4))
    ax.text(t, y_pos, msg, rotation=90, fontsize=6.5,
            va="top", ha="right", color="gray", alpha=0.55)

t_right = max(t_last + timedelta(hours=2), df["time"].max() + padding)
ax.set_xlim(t_first - padding, t_right)
ax.set_ylabel("Jaccard similarity (%)")
ax.set_xlabel("Time")
ax.set_title("Visual similarity over time")
ax.legend()
fig.autofmt_xdate()
plt.tight_layout()
plt.show()
