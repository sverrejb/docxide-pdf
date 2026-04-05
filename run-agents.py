#!/usr/bin/env python3
"""Launch parallel Claude agents in worktrees to improve test case scores.

Usage:
    python3 run-agents.py case42 case43 case44       # specific cases
    python3 run-agents.py --worst 3                  # auto-pick 3 worst from new/
    python3 run-agents.py --worst 6 --workers 3      # 6 cases across 3 workers
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from textwrap import dedent

REPO_ROOT = Path(__file__).resolve().parent
FIXTURES_DIR = REPO_ROOT / "tests" / "fixtures"
BASELINES_PATH = REPO_ROOT / "tests" / "baselines.json"
SKIPLIST_PATH = FIXTURES_DIR / "SKIPLIST"

# Gitignored dirs to symlink into worktrees
GITIGNORED_DIRS = [
    "tests/output",
    "tests/fixtures/scraped",
    "tools/target",
    "target",
]


# ── Helpers ──────────────────────────────────────────────────────────────────


def load_skiplist() -> set[str]:
    if not SKIPLIST_PATH.exists():
        return set()
    skips = set()
    for line in SKIPLIST_PATH.read_text().splitlines():
        line = line.split("#")[0].strip()
        if line:
            skips.add(line)
    return skips


def load_baselines() -> dict:
    with open(BASELINES_PATH) as f:
        return json.load(f)


def baseline_key_for(dirname: str) -> str:
    """Baselines truncate long names to 20 chars + '..'"""
    if len(dirname) > 20:
        return f"new/{dirname[:20]}.."
    return f"new/{dirname}"


def pick_worst_cases(n: int) -> list[str]:
    """Select the N worst-scoring cases from new/ group."""
    baselines = load_baselines()
    skips = load_skiplist()

    new_dir = FIXTURES_DIR / "new"
    if not new_dir.is_dir():
        print("Error: tests/fixtures/new/ does not exist", file=sys.stderr)
        sys.exit(1)

    real_dirs = [d.name for d in new_dir.iterdir() if d.is_dir()]
    key_to_dir = {baseline_key_for(d): d for d in real_dirs}

    scored = []
    for key, v in baselines.items():
        if not key.startswith("new/"):
            continue
        real_name = key_to_dir.get(key, key.split("/")[-1])
        if real_name in skips or "new" in skips:
            continue
        avg_score = (v.get("jaccard", 0) + v.get("ssim", 0)) / 2
        scored.append((avg_score, real_name))

    scored.sort()
    return [name for _, name in scored[:n]]


def resolve_case_path(name: str) -> str:
    """Resolve case name to group/case path (e.g. 'case42' -> 'cases/case42')."""
    for group_dir in sorted(FIXTURES_DIR.iterdir()):
        if group_dir.is_dir() and (group_dir / name).is_dir():
            return f"{group_dir.name}/{name}"
    print(f"Error: fixture '{name}' not found in any group", file=sys.stderr)
    sys.exit(1)


def get_scores(case_path: str) -> str:
    """Look up current scores from baselines.json."""
    baselines = load_baselines()
    if case_path in baselines:
        v = baselines[case_path]
        return f"Jaccard: {v.get('jaccard', 'N/A')}, SSIM: {v.get('ssim', 'N/A')}, Text boundary: {v.get('text_boundary', 'N/A')}"
    return "No baseline scores found"


def count_annotations(case_name: str) -> int:
    ann_path = REPO_ROOT / "tests" / "output" / "annotations.json"
    if not ann_path.exists():
        return 0
    try:
        with open(ann_path) as f:
            data = json.load(f)
        return sum(
            1
            for a in data
            if a.get("case") == case_name and not a.get("fixed", False)
        )
    except Exception:
        return 0


# ── Worktree management ─────────────────────────────────────────────────────


def run_git(*args, check=True, **kwargs) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["git", "-C", str(REPO_ROOT), *args],
        capture_output=True,
        text=True,
        check=check,
        **kwargs,
    )


def setup_worktree(case_name: str) -> Path:
    wt_name = f"agent-{case_name}"
    wt_path = REPO_ROOT / ".worktrees" / wt_name
    branch = f"agent/{wt_name}"

    # Try creating; if branch/worktree already exists, clean up and retry
    result = run_git("worktree", "add", "-b", branch, str(wt_path), "HEAD", check=False)
    if result.returncode != 0:
        run_git("worktree", "remove", "--force", str(wt_path), check=False)
        run_git("branch", "-D", branch, check=False)
        run_git("worktree", "add", "-b", branch, str(wt_path), "HEAD")

    # Symlink gitignored directories
    for rel in GITIGNORED_DIRS:
        src = REPO_ROOT / rel
        dst = wt_path / rel
        if src.is_dir():
            dst.parent.mkdir(parents=True, exist_ok=True)
            if dst.exists() or dst.is_symlink():
                if dst.is_symlink() or dst.is_file():
                    dst.unlink()
                else:
                    import shutil

                    shutil.rmtree(dst)
            dst.symlink_to(src)

    return wt_path


# ── Prompt ───────────────────────────────────────────────────────────────────


def build_prompt(case_name: str, case_path: str, progress_file: Path, logs_dir: Path) -> str:
    scores = get_scores(case_path)
    ann_count = count_annotations(case_name)

    annotations_section = ""
    if ann_count > 0:
        annotations_section = f"""
4. **Read annotations.** There are {ann_count} unfixed annotations for this case in tests/output/annotations.json. Read them first — they contain precise coordinates and descriptions of rendering issues.
"""

    return dedent(f"""\
        You are working on the docxside-pdf project — a Rust library that converts DOCX files to PDF.
        Your task is to improve the rendering quality for a specific test fixture: **{case_name}** (path: tests/fixtures/{case_path}/).

        ## Current scores
        {scores}

        ## Your goal
        Improve the Jaccard similarity and/or SSIM score for this case. Even small improvements (1-5%) are valuable. Focus on the most impactful issues first.

        ## Progress logging
        After each significant action (investigation finding, code change, test run), append a timestamped entry to your progress file:
          echo "$(date '+%H:%M:%S') — <what you did and what happened>" >> "{progress_file}"

        Do this throughout your work so I can monitor progress.

        ## Workflow
        1. **Investigate first.** Run the test for just this case to confirm the baseline:
           DOCXIDE_CASE={case_name} cargo test visual_comparison -- --nocapture
           Log the starting scores.

        2. **Inspect the fixture.** Use `./tools/target/debug/docx-inspect tests/fixtures/{case_path}/input.docx` to understand what DOCX features are used. Dump specific XML files to see the structure.

        3. **Compare output.** After running the test, look at the diff images in tests/output/{case_path}/diff/ — blue pixels = reference only, red = generated only. The reference screenshots are in tests/output/{case_path}/reference/ and generated in tests/output/{case_path}/generated/.
        {annotations_section}
        4. **Identify the root cause.** What's the biggest visual difference? Is it a missing feature, wrong spacing, missing font, incorrect layout?

        5. **Make targeted fixes** in the Rust source code. Focus on fixes that help this case without breaking others. After each change, run:
           DOCXIDE_CASE={case_name} cargo test visual_comparison -- --nocapture
           to see if scores improved.

        6. **Check for regressions.** Run the full test suite before finalizing:
           cargo test visual_comparison -- --nocapture

        7. **Accept baselines** for all changed scores:
           ./tools/target/debug/accept-baselines

        ## Finalization — MANDATORY

        When you are done (whether you improved scores or not), you MUST write an outcome file.
        This file is machine-read by the orchestrator script to decide whether to auto-merge your work.

        **Write this file as your very last action** to the path: `{logs_dir / (case_name + ".outcome.json")}`

        The JSON must have this structure:
        ```json
        {{
          "case": "{case_name}",
          "case_path": "{case_path}",
          "improved": true,
          "target_before": {{"jaccard": 0.0, "ssim": 0.0}},
          "target_after": {{"jaccard": 0.0, "ssim": 0.0}},
          "regressions": [],
          "summary": "Description of what you changed and why"
        }}
        ```

        Rules for the outcome:
        - **"improved": true** only if Jaccard OR SSIM increased by at least 0.5% (0.005) on the target case
        - **"regressions"** lists ANY other case where Jaccard or SSIM dropped by more than 2% (0.02) compared to the committed baselines.json. Each entry: {{"case": "name", "metric": "jaccard|ssim", "before": 0.0, "after": 0.0}}. Leave as empty array [] if no regressions.
        - If you made no code changes (e.g. the issue is font-related and unfixable), set improved to false and explain in summary.
        - **Commit your changes** (code + updated baselines.json) before writing the outcome file. Use a descriptive commit message. End every commit message with this trailer on its own line:
          Automated-by: run-agents.py ({case_name})

        The orchestrator will:
        - **Auto-merge** your branch into main if improved=true AND regressions=[]
        - **Flag for manual review** if there are regressions or if the outcome is ambiguous

        ## Important rules
        - Do NOT modify test fixtures or reference PDFs
        - Do NOT modify the test harness or scoring thresholds
        - Keep changes minimal and focused — don't refactor unrelated code
        - If the case fails due to missing fonts, log that finding and move on to layout/rendering issues instead
        - The DOCX spec can be queried via the local RAG tool (mcp__local-rag__query_documents) if you need to look up XML element semantics

        ## Key environment
        - Filter to single case: DOCXIDE_CASE={case_name} cargo test visual_comparison -- --nocapture
        - Inspect DOCX: ./tools/target/debug/docx-inspect tests/fixtures/{case_path}/input.docx [path]
        - Analysis: ./tools/target/debug/analyze-fixtures --grep "pattern"
    """)


# ── Agent launcher ───────────────────────────────────────────────────────────


def launch_agent(
    case_name: str,
    case_path: str,
    logs_dir: Path,
    model: str,
    max_turns: int | None,
    permission_mode: str,
) -> dict:
    """Run a single agent in a worktree. Called in a subprocess via ProcessPoolExecutor."""
    progress_file = logs_dir / f"{case_name}.log"
    outcome_file = logs_dir / f"{case_name}.outcome.json"

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    progress_file.write_text(f"{timestamp} — Agent started for {case_name} ({case_path})\n")

    try:
        wt_path = setup_worktree(case_name)
    except Exception as e:
        progress_file.write_text(f"{timestamp} — Failed to setup worktree: {e}\n")
        return {"case": case_name, "error": f"worktree setup failed: {e}"}

    prompt = build_prompt(case_name, case_path, progress_file, logs_dir)

    cmd = [
        "claude", "-p",
        "--model", model,
        "--effort", "high",
        "--permission-mode", permission_mode,
        prompt,
    ]
    if max_turns is not None:
        cmd.extend(["--max-turns", str(max_turns)])

    print(f"  [{case_name}] Started in {wt_path}")

    with open(progress_file, "a") as log:
        result = subprocess.run(
            cmd,
            cwd=str(wt_path),
            stdout=log,
            stderr=subprocess.STDOUT,
        )

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(progress_file, "a") as log:
        log.write(f"{timestamp} — Agent finished (exit code {result.returncode})\n")

    print(f"  [{case_name}] Finished (exit code {result.returncode})")

    return {
        "case": case_name,
        "case_path": case_path,
        "wt_path": str(wt_path),
        "exit_code": result.returncode,
        "outcome_file": str(outcome_file),
    }


# ── Outcome processing ──────────────────────────────────────────────────────


def parse_outcome(outcome_file: str) -> dict | None:
    path = Path(outcome_file)
    if not path.exists():
        return None
    try:
        with open(path) as f:
            data = json.load(f)
        reg = data.get("regressions", [])
        return {
            "improved": bool(data.get("improved")),
            "regressions": reg if isinstance(reg, list) else [],
            "summary": data.get("summary", "No summary"),
            "target_before": data.get("target_before", {}) if isinstance(data.get("target_before"), dict) else {},
            "target_after": data.get("target_after", {}) if isinstance(data.get("target_after"), dict) else {},
        }
    except Exception as e:
        return {"error": str(e)}


def try_merge(case_name: str, wt_path: str) -> tuple[bool, str]:
    """Attempt to merge the agent branch into main. Returns (success, message)."""
    branch = f"agent/agent-{case_name}"

    # Check commits ahead
    result = subprocess.run(
        ["git", "-C", wt_path, "rev-list", "HEAD", "^main", "--count"],
        capture_output=True, text=True, check=False,
    )
    ahead = int(result.stdout.strip()) if result.returncode == 0 else 0
    if ahead == 0:
        return False, "No commits to merge"

    # Merge
    result = subprocess.run(
        ["git", "-C", str(REPO_ROOT), "merge", "--no-edit", branch],
        capture_output=True, text=True, check=False,
    )
    if result.returncode != 0:
        # Abort failed merge
        subprocess.run(
            ["git", "-C", str(REPO_ROOT), "merge", "--abort"],
            capture_output=True, check=False,
        )
        return False, f"Merge conflict:\n{result.stdout}\n{result.stderr}"

    # Cleanup worktree and branch
    subprocess.run(
        ["git", "-C", str(REPO_ROOT), "worktree", "remove", wt_path, "--force"],
        capture_output=True, check=False,
    )
    subprocess.run(
        ["git", "-C", str(REPO_ROOT), "branch", "-D", branch],
        capture_output=True, check=False,
    )

    return True, f"Merged {ahead} commit(s) from {branch}"


def process_results(results: list[dict], logs_dir: Path):
    """Process agent results: auto-merge or flag for review."""
    print("\n" + "=" * 60)
    print("Processing outcomes...")
    print("=" * 60 + "\n")

    merged = []
    flagged = []
    no_outcome = []

    for r in results:
        case_name = r["case"]
        case_path = r.get("case_path", "?")
        wt_path = r.get("wt_path", "?")

        print(f"── {case_name} ({case_path}) ──")

        if "error" in r:
            print(f"  Error: {r['error']}")
            flagged.append((case_name, r["error"]))
            print()
            continue

        outcome = parse_outcome(r["outcome_file"])
        if outcome is None:
            print(f"  No outcome file — flagging for manual review")
            print(f"  Worktree: {wt_path}")
            print(f"  Log: {logs_dir / (case_name + '.log')}")
            no_outcome.append(case_name)
            print()
            continue

        if "error" in outcome:
            print(f"  Failed to parse outcome: {outcome['error']}")
            flagged.append((case_name, f"bad outcome JSON: {outcome['error']}"))
            print()
            continue

        before = outcome["target_before"]
        after = outcome["target_after"]
        regs = outcome["regressions"]

        print(f"  Summary: {outcome['summary']}")
        print(f"  Jaccard: {before.get('jaccard', '?')} -> {after.get('jaccard', '?')}")
        print(f"  SSIM:    {before.get('ssim', '?')} -> {after.get('ssim', '?')}")

        if outcome["improved"] and len(regs) == 0:
            print(f"  Improved with no regressions — auto-merging into main")
            ok, msg = try_merge(case_name, wt_path)
            if ok:
                merged.append(case_name)
                print(f"  Merged: {msg}")
            else:
                flagged.append((case_name, f"merge failed: {msg}"))
                print(f"  Merge failed: {msg}")
                print(f"  Worktree: {wt_path}")

        elif outcome["improved"] and len(regs) > 0:
            print(f"  Improved but has {len(regs)} regression(s) — flagging for manual review")
            for reg in regs:
                print(f"    {reg.get('case', '?')}: {reg.get('metric', '?')} {reg.get('before', '?')} -> {reg.get('after', '?')}")
            flagged.append((case_name, "regressions"))
            print(f"  Worktree: {wt_path}")
            print(f"  Review:   cd {wt_path} && git log --oneline main..HEAD")

        else:
            print(f"  No meaningful improvement")
            flagged.append((case_name, "no improvement"))
            print(f"  Worktree: {wt_path}")

        print()

    # Final summary
    print("=" * 60)
    print("SUMMARY")
    print("=" * 60 + "\n")

    if merged:
        print(f"Auto-merged ({len(merged)}):")
        for c in merged:
            print(f"  + {c}")
        print()

    if flagged:
        print(f"Flagged for manual review ({len(flagged)}):")
        for c, reason in flagged:
            print(f"  ! {c} ({reason})")
        print()
        print("To review a flagged case:")
        print(f"  cd {REPO_ROOT}/.worktrees/agent-<case>")
        print(f"  git log --oneline main..HEAD")
        print(f"  git diff main")
        print(f"  # If happy: git -C {REPO_ROOT} merge agent/agent-<case>")
        print()

    if no_outcome:
        print(f"No outcome file ({len(no_outcome)}):")
        for c in no_outcome:
            print(f"  ? {c} — check {logs_dir / (c + '.log')}")
        print()

    print(f"Logs: {logs_dir}/")
    print()
    print("To clean up ALL remaining worktrees:")
    print(f"  for d in {REPO_ROOT}/.worktrees/agent-*; do")
    print(f'    name=$(basename "$d"); git -C {REPO_ROOT} worktree remove "$d" --force; git -C {REPO_ROOT} branch -D "agent/$name" 2>/dev/null')
    print(f"  done")


# ── Main ─────────────────────────────────────────────────────────────────────


def main():
    parser = argparse.ArgumentParser(description="Launch parallel Claude agents to improve test cases")
    parser.add_argument("cases", nargs="*", help="Case names to work on")
    parser.add_argument("--worst", type=int, default=0, help="Auto-select N worst-scoring cases from new/")
    parser.add_argument("--workers", type=int, default=3, help="Number of parallel agents (default: 3)")
    parser.add_argument("--model", default="opus", help="Claude model (default: opus)")
    parser.add_argument("--max-turns", type=int, default=None, help="Max conversation turns per agent")
    parser.add_argument("--permission", default="auto", help="Permission mode (default: auto)")
    args = parser.parse_args()

    cases = list(args.cases)
    if args.worst > 0:
        worst = pick_worst_cases(args.worst)
        cases.extend(worst)
        print(f"Auto-selected worst {args.worst} cases: {' '.join(worst)}")

    if not cases:
        parser.error("No cases specified. Use --worst N or pass case names.")

    # Resolve to group/case paths
    case_paths = {name: resolve_case_path(name) for name in cases}

    # Create logs directory
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    logs_dir = REPO_ROOT / "logs" / f"agents-{timestamp}"
    logs_dir.mkdir(parents=True, exist_ok=True)
    print(f"Logs: {logs_dir}")

    print(f"\nLaunching {len(cases)} cases across {args.workers} workers...")
    print(f"Cases: {' '.join(cases)}\n")

    # Launch agents in parallel with worker pool
    results = []
    with ProcessPoolExecutor(max_workers=args.workers) as pool:
        futures = {
            pool.submit(
                launch_agent,
                name,
                case_paths[name],
                logs_dir,
                args.model,
                args.max_turns,
                args.permission,
            ): name
            for name in cases
        }
        for future in as_completed(futures):
            name = futures[future]
            try:
                results.append(future.result())
            except Exception as e:
                print(f"  [{name}] Exception: {e}")
                results.append({"case": name, "error": str(e)})

    # Sort results back to original case order
    order = {name: i for i, name in enumerate(cases)}
    results.sort(key=lambda r: order.get(r["case"], 999))

    print(f"\nMonitor progress: tail -f {logs_dir}/*.log")

    process_results(results, logs_dir)


if __name__ == "__main__":
    main()
