#!/usr/bin/env python3
"""Launch parallel Claude agents in worktrees to improve test case scores.

Usage:
    python3 run-agents.py case42 case43 case44       # specific cases
    python3 run-agents.py --worst 3                  # auto-pick 3 worst from new/
    python3 run-agents.py --worst 6 --workers 3      # 6 cases across 3 workers
    python3 run-agents.py --annotations 4            # auto-pick 4 random annotated cases
    python3 run-agents.py --discover 3               # download 3 new DOCX, name, convert, then work on them
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import shutil
import subprocess
import sys
import threading
import time
import uuid
import zipfile
from concurrent.futures import ProcessPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from textwrap import dedent
from xml.etree import ElementTree as ET

REPO_ROOT = Path(__file__).resolve().parent
FIXTURES_DIR = REPO_ROOT / "tests" / "fixtures"
BASELINES_PATH = REPO_ROOT / "tests" / "baselines.json"
SKIPLIST_PATH = FIXTURES_DIR / "SKIPLIST"
CASE_HISTORY_DIR = REPO_ROOT / "logs" / "case-history"
DOWNLOAD_DIR = REPO_ROOT / "downloads"
MANIFEST_PATH = REPO_ROOT.parent / "docx-corpus" / "manifest.txt"
NEW_DIR = FIXTURES_DIR / "new"

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Gitignored dirs to symlink into worktrees (shared, read-heavy)
SYMLINK_DIRS = [
    "tests/fixtures/scraped",
    "tools/target",
    "target",
]

# Files to copy from tests/output/ into the worktree's own output dir
OUTPUT_SEED_FILES = [
    "annotations.json",
    "latest_scores.json",
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
    """Baselines truncate long names to 16 chars + '..'"""
    if len(dirname) > 16:
        return f"new/{dirname[:16]}.."
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
        ssim = v.get("ssim", 0)
        scored.append((ssim, real_name))

    scored.sort()
    return [name for _, name in scored[:n]]


def pick_annotated_cases(n: int) -> list[str]:
    """Select N random unique cases that have unfixed annotations."""
    ann_path = REPO_ROOT / "tests" / "output" / "annotations.json"
    if not ann_path.exists():
        print("Error: tests/output/annotations.json does not exist", file=sys.stderr)
        sys.exit(1)

    with open(ann_path) as f:
        data = json.load(f)

    annotations = data.get("annotations", data) if isinstance(data, dict) else data
    unfixed = [a for a in annotations if not a.get("fixed", False)]
    if not unfixed:
        print("Error: no unfixed annotations found", file=sys.stderr)
        sys.exit(1)

    # Extract unique case names (strip group prefix — resolve_case_path will re-resolve)
    cases = set()
    for a in unfixed:
        case = a.get("case", "")
        # Annotation case field may be "group/name" or bare "name"
        name = case.split("/")[-1] if "/" in case else case
        if name:
            cases.add(name)

    cases = sorted(cases)
    if n >= len(cases):
        return cases

    return sorted(random.sample(cases, n))


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


def get_annotations(case_name: str, case_path: str) -> list[dict]:
    """Get unfixed annotations for a case. Matches on both case_path and case_name."""
    ann_path = REPO_ROOT / "tests" / "output" / "annotations.json"
    if not ann_path.exists():
        return []
    try:
        with open(ann_path) as f:
            data = json.load(f)
        annotations = data.get("annotations", data) if isinstance(data, dict) else data
        return [
            a for a in annotations
            if not a.get("fixed", False)
            and a.get("case") in (case_name, case_path, f"{case_path}")
        ]
    except Exception:
        return []


def format_annotations(annotations: list[dict]) -> str:
    """Format annotations into readable text for the agent prompt."""
    if not annotations:
        return ""
    lines = []
    for a in annotations:
        page = a.get("page", "?")
        source = a.get("source", "?")
        x = a.get("x_pt", 0)
        y = a.get("y_pt", 0)
        note = a.get("note", "")
        lines.append(f"  - Page {page} ({source}) at ({x:.1f}, {y:.1f}pt): {note}")
    return "\n".join(lines)


MAX_HISTORY_CHARS = 4000  # ~1k tokens — enough for context, not a context hog


def get_case_history(case_name: str) -> str:
    """Read persistent history for a case from previous agent runs."""
    history_file = CASE_HISTORY_DIR / f"{case_name}.md"
    if not history_file.exists():
        return ""
    text = history_file.read_text()
    if len(text) > MAX_HISTORY_CHARS:
        # Keep the most recent entries (at the end of the file)
        text = "... (earlier history truncated) ...\n" + text[-MAX_HISTORY_CHARS:]
    return text


def history_file_for(case_name: str) -> Path:
    """Return the path to the persistent history file for a case."""
    CASE_HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    return CASE_HISTORY_DIR / f"{case_name}.md"


# ── Discovery helpers ──────────────────────────────────────────────────────


def collect_existing_content_hashes() -> set[str]:
    """SHA-256 hash every input.docx across all fixture groups."""
    hashes: set[str] = set()
    for group_dir in FIXTURES_DIR.iterdir():
        if group_dir.is_dir():
            for case_dir in group_dir.iterdir():
                docx = case_dir / "input.docx"
                if docx.exists():
                    hashes.add(hashlib.sha256(docx.read_bytes()).hexdigest())
    return hashes


def collect_existing_names() -> set[str]:
    """All directory names across all fixture groups."""
    names: set[str] = set()
    for group_dir in FIXTURES_DIR.iterdir():
        if group_dir.is_dir():
            for case_dir in group_dir.iterdir():
                if case_dir.is_dir():
                    names.add(case_dir.name)
    return names


def extract_text_snippet(docx_path: Path, max_chars: int = 500) -> str:
    """Extract raw text from a DOCX for naming purposes."""
    try:
        with zipfile.ZipFile(docx_path) as zf:
            if "word/document.xml" not in zf.namelist():
                return ""
            xml = zf.read("word/document.xml")
        root = ET.fromstring(xml)
        texts = []
        for t in root.iter(f"{{{WML_NS}}}t"):
            if t.text:
                texts.append(t.text)
            if sum(len(s) for s in texts) > max_chars:
                break
        return " ".join(texts)[:max_chars]
    except Exception:
        return ""


def name_files_with_claude(file_snippets: dict[str, str], existing_names: set[str]) -> dict[str, str]:
    """Call claude -p once to name all files. Returns {hash: name}."""
    existing_list = ", ".join(sorted(existing_names)) if existing_names else "(none)"

    entries = []
    for h, text in file_snippets.items():
        preview = text[:300].replace("\n", " ").strip()
        entries.append(f"- **{h}**: {preview}")
    file_list = "\n".join(entries)

    prompt = f"""\
I have {len(file_snippets)} DOCX files that need descriptive English snake_case names.
The names should describe what the document is about (e.g. "czech_health_statement_form", "russian_volunteerism_essay", "who_prescribing_patterns_table").

Rules:
- English snake_case, 3-5 words, max 40 chars
- Describe the document's topic/purpose, not its language
- If the text is in a non-English language, translate the topic to English
- Each name must be unique and not collide with existing names
- No generic names like "document_1" or "test_file"

Existing names (avoid these): {existing_list}

Files to name (hash: text preview):
{file_list}

Respond with ONLY a JSON object mapping hash to name, nothing else. Example:
{{"abc123": "italian_budget_report", "def456": "korean_school_schedule"}}"""

    print("Asking Claude to name files...")
    result = subprocess.run(
        ["claude", "-p", "--model", "haiku", "--output-format", "json", prompt],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(f"Error: Claude naming failed: {result.stderr}", file=sys.stderr)
        return {}

    try:
        response = json.loads(result.stdout)
        if "result" in response and isinstance(response["result"], str):
            inner = response["result"]
            start = inner.find("{")
            end = inner.rfind("}") + 1
            if start >= 0 and end > start:
                return json.loads(inner[start:end])
        elif all(isinstance(v, str) for v in response.values()):
            return response
    except (json.JSONDecodeError, AttributeError):
        try:
            text = result.stdout
            start = text.find("{")
            end = text.rfind("}") + 1
            if start >= 0 and end > start:
                return json.loads(text[start:end])
        except json.JSONDecodeError:
            pass

    print(f"Error: could not parse Claude naming response: {result.stdout[:500]}", file=sys.stderr)
    return {}


DISMISS_SCRIPT = """
    tell application "System Events"
        tell process "Microsoft Word"
            if exists (button "Yes" of window 1) then
                click button "Yes" of window 1
            else if exists (button "OK" of window 1) then
                click button "OK" of window 1
            end if
        end tell
    end tell
"""


def _dialog_watcher(stop_event: threading.Event) -> None:
    while not stop_event.is_set():
        subprocess.run(["osascript", "-e", DISMISS_SCRIPT], capture_output=True)
        time.sleep(0.5)


def convert_docx_to_reference_pdf(docx_path: Path, pdf_path: Path) -> bool:
    """Convert a DOCX to reference PDF via MS Word. Returns True on success."""
    staging = Path.home() / "Documents" / f"_docx_convert_{uuid.uuid4().hex}"
    staging.mkdir()
    try:
        tmp_docx = staging / "input.docx"
        tmp_pdf = staging / "input.pdf"
        shutil.copy2(docx_path, tmp_docx)
        subprocess.run(["xattr", "-d", "com.apple.quarantine", str(tmp_docx)], capture_output=True)

        script = f"""
            tell application "Microsoft Word"
                set display alerts to -2
                open POSIX file "{tmp_docx.resolve()}"
                delay 2
                set theDoc to document 1
                save as theDoc file name "{tmp_pdf.resolve()}" file format format PDF
                close theDoc saving no
                set display alerts to 0
            end tell
        """
        stop_event = threading.Event()
        watcher = threading.Thread(target=_dialog_watcher, args=(stop_event,), daemon=True)
        watcher.start()
        result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        stop_event.set()

        if result.returncode != 0:
            print(f"  Word conversion failed: {result.stderr.strip()}", file=sys.stderr)
            return False
        if not tmp_pdf.exists():
            print("  Word conversion produced no PDF", file=sys.stderr)
            return False

        shutil.move(str(tmp_pdf), pdf_path)
        return True
    finally:
        shutil.rmtree(staging, ignore_errors=True)


def discover_new_cases(
    count: int,
    min_size: int = 10_000,
    max_size: int = 500_000,
) -> list[str]:
    """Download, deduplicate, name, and convert new DOCX files. Returns list of case names."""
    if not MANIFEST_PATH.exists():
        print(f"Error: manifest not found at {MANIFEST_PATH}", file=sys.stderr)
        sys.exit(1)

    all_hashes = [h.strip() for h in MANIFEST_PATH.read_text().splitlines() if h.strip()]
    existing_content_hashes = collect_existing_content_hashes()
    existing_names = collect_existing_names()

    # Skip manifest hashes we've already downloaded
    existing_manifest_hashes: set[str] = set()
    if DOWNLOAD_DIR.exists():
        for f in DOWNLOAD_DIR.glob("*.docx"):
            existing_manifest_hashes.add(f.stem)

    available = [h for h in all_hashes if h not in existing_manifest_hashes]
    random.shuffle(available)
    print(f"Manifest: {len(all_hashes)} total, {len(existing_content_hashes)} existing fixtures, {len(available)} available")

    DOWNLOAD_DIR.mkdir(exist_ok=True)
    NEW_DIR.mkdir(parents=True, exist_ok=True)

    # Phase 1: Download candidates (over-fetch to account for filtering)
    candidates: list[tuple[str, Path]] = []
    target = count * 3

    for h in available:
        if len(candidates) >= target:
            break

        dest = DOWNLOAD_DIR / f"{h}.docx"
        if not dest.exists():
            print(f"  Downloading {h[:16]}...")
            result = subprocess.run(
                ["curl", "-sf", "-o", str(dest), f"https://docxcorp.us/documents/{h}.docx"],
                capture_output=True,
            )
            if result.returncode != 0:
                continue

        size = dest.stat().st_size
        if size < min_size or size > max_size:
            continue
        if not zipfile.is_zipfile(dest):
            continue

        content_hash = hashlib.sha256(dest.read_bytes()).hexdigest()
        if content_hash in existing_content_hashes:
            print(f"  {h[:16]}: duplicate content, skipping")
            continue
        existing_content_hashes.add(content_hash)

        text = extract_text_snippet(dest)
        if len(text.strip()) < 30:
            continue

        candidates.append((h, dest))

    print(f"Downloaded {len(candidates)} valid candidates for {count} slots")
    candidates = candidates[:count]

    if not candidates:
        print("Error: no valid candidates found", file=sys.stderr)
        sys.exit(1)

    # Phase 2: Name with Claude
    snippets = {h: extract_text_snippet(path) for h, path in candidates}
    names = name_files_with_claude(snippets, existing_names)
    if not names:
        print("Error: Claude naming returned no names", file=sys.stderr)
        sys.exit(1)

    used_names: set[str] = set(existing_names)
    hash_to_name: dict[str, str] = {}
    for h, _ in candidates:
        name = names.get(h)
        if not name or name in used_names:
            print(f"  Warning: no valid name for {h[:16]} (got {name!r}), skipping")
            continue
        name = name.strip().lower().replace(" ", "_").replace("-", "_")
        name = "".join(c for c in name if c.isalnum() or c == "_").strip("_")
        if not name or name in used_names:
            continue
        hash_to_name[h] = name
        used_names.add(name)

    print(f"Named {len(hash_to_name)} files:")
    for h, name in hash_to_name.items():
        print(f"  {h[:16]} -> {name}")

    # Phase 3: Create fixtures and convert to PDF
    pending = [(h, p) for h, p in candidates if hash_to_name.get(h)]
    print(f"\nReady to convert {len(pending)} file(s) to PDF via MS Word.")
    print("Word will open and may show dialogs that need to be dismissed.")
    try:
        input("Press Enter when you're ready to start (Ctrl+C to abort)... ")
    except (KeyboardInterrupt, EOFError):
        print("\nAborted before PDF conversion.", file=sys.stderr)
        sys.exit(1)

    created: list[str] = []
    for h, docx_path in candidates:
        name = hash_to_name.get(h)
        if not name:
            continue

        case_dir = NEW_DIR / name
        case_dir.mkdir(exist_ok=True)
        shutil.copy2(docx_path, case_dir / "input.docx")

        print(f"  Converting {name} to PDF via Word...")
        if convert_docx_to_reference_pdf(case_dir / "input.docx", case_dir / "reference.pdf"):
            print(f"  {name} — OK")
            created.append(name)
        else:
            print(f"  {name} — PDF conversion failed, removing fixture")
            shutil.rmtree(case_dir)

    print(f"Created {len(created)} new fixtures in {NEW_DIR}")

    # Commit new fixtures to main so worktrees (branched from HEAD) include them
    if created:
        for name in created:
            subprocess.run(
                ["git", "-C", str(REPO_ROOT), "add", str(NEW_DIR / name)],
                capture_output=True, check=False,
            )
        names_str = ", ".join(created)
        subprocess.run(
            ["git", "-C", str(REPO_ROOT), "commit", "-m", f"Add discovered fixtures: {names_str}"],
            capture_output=True, check=False,
        )
        print(f"Committed {len(created)} new fixtures to main")

    return created


# ── Worktree management ─────────────────────────────────────────────────────


def run_git(*args, check=True, **kwargs) -> subprocess.CompletedProcess:
    return subprocess.run(
        ["git", "-C", str(REPO_ROOT), *args],
        capture_output=True,
        text=True,
        check=check,
        **kwargs,
    )


def setup_worktree(case_name: str, resume: bool = False) -> Path:
    wt_name = f"agent-{case_name}"
    wt_path = REPO_ROOT / ".worktrees" / wt_name
    branch = f"agent/{wt_name}"

    if resume:
        if not wt_path.is_dir():
            raise RuntimeError(f"resume requested but worktree not found: {wt_path}")
        return wt_path

    # Try creating; if branch/worktree already exists, clean up and retry
    result = run_git("worktree", "add", "-b", branch, str(wt_path), "HEAD", check=False)
    if result.returncode != 0:
        run_git("worktree", "remove", "--force", str(wt_path), check=False)
        run_git("branch", "-D", branch, check=False)
        run_git("worktree", "add", "-b", branch, str(wt_path), "HEAD")

    import shutil

    # Symlink shared gitignored directories
    for rel in SYMLINK_DIRS:
        src = REPO_ROOT / rel
        dst = wt_path / rel
        if src.is_dir():
            dst.parent.mkdir(parents=True, exist_ok=True)
            if dst.exists() or dst.is_symlink():
                if dst.is_symlink() or dst.is_file():
                    dst.unlink()
                else:
                    shutil.rmtree(dst)
            dst.symlink_to(src)

    # Create a private tests/output/ so parallel agents don't clobber each other
    wt_output = wt_path / "tests" / "output"
    wt_output.mkdir(parents=True, exist_ok=True)
    main_output = REPO_ROOT / "tests" / "output"
    for fname in OUTPUT_SEED_FILES:
        src = main_output / fname
        if src.exists():
            shutil.copy2(str(src), str(wt_output / fname))

    return wt_path


# ── Prompt ───────────────────────────────────────────────────────────────────


def build_prompt(case_name: str, case_path: str, progress_file: Path, logs_dir: Path) -> str:
    scores = get_scores(case_path)
    annotations = get_annotations(case_name, case_path)
    history = get_case_history(case_name)
    history_path = history_file_for(case_name)

    annotations_section = ""
    if annotations:
        formatted = format_annotations(annotations)
        annotations_section = (
            f"\n        **Human-annotated issues ({len(annotations)}):** These are manually identified rendering problems"
            f" for this case. Prioritize fixing these — they describe exactly what's wrong and where.\n\n"
            f"{formatted}\n"
        )

    history_section = ""
    if history:
        history_section = (
            "\n        ## History from previous runs\n"
            "        The following is a log from previous agent runs on this case. Read it carefully —"
            " it documents what was tried, what worked, what failed, and what was learned."
            " Do NOT repeat failed approaches.\n\n"
            f"        ```\n{history}\n        ```\n"
        )

    return dedent(f"""\
        You are working on the docxside-pdf project — a Rust library that converts DOCX files to PDF.
        Your task is to improve the rendering quality for a specific test fixture: **{case_name}** (path: tests/fixtures/{case_path}/).

        ## Current scores
        {scores}
        {history_section}
        ## Your goal
        Improve the Jaccard similarity and/or SSIM score for this case. Even small improvements (1-5%) are valuable. Focus on the most impactful issues first.

        ## PRIME DIRECTIVE — what to focus on
        Focus on **bugs, missing features, and things we are not rendering** — e.g. missing borders, backgrounds, images, shapes, text effects, paragraph properties, or entire content blocks that are absent from our output.

        Do NOT work on:
        - Word/character spacing or kerning precision
        - Line height or line spacing drift
        - Vertical or horizontal positional drift/accumulation
        - Any x/y micro-alignment issues

        These spacing/drift issues are known and tracked separately. Your time is best spent on things that are completely wrong or missing, not on fine-tuning positions.

        ## Progress logging
        After each significant action (investigation finding, code change, test run), append a timestamped entry to your progress file:
          echo "$(date '+%H:%M:%S') — <what you did and what happened>" >> "{progress_file}"

        Do this throughout your work so I can monitor progress.

        ## Workflow
        1. **Investigate first.** Run the test for just this case to confirm the baseline:
           ./tools/run-tests.sh --case {case_name} --test visual_comparison
           Log the starting scores.

        2. **Inspect the fixture.** Use `./tools/target/debug/docx-inspect tests/fixtures/{case_path}/input.docx` to list ZIP entries. Then dump only the specific XML files you need (e.g. `word/document.xml`, `word/styles.xml`). Do NOT dump everything at once — large XML files waste context.

        3. **Compare output.** After running the test, look at the diff images in tests/output/{case_path}/diff/ — blue pixels = reference only, red = generated only. The reference screenshots are in tests/output/{case_path}/reference/ and generated in tests/output/{case_path}/generated/.
        {annotations_section}
        4. **Identify the root cause.** What's the biggest visual difference? Is it a missing feature, wrong spacing, missing font, incorrect layout?

        5. **Make targeted fixes** in the Rust source code. Focus on fixes that help this case without breaking others. After each change, run:
           ./tools/run-tests.sh --case {case_name} --test visual_comparison
           to see if scores improved.

        6. **Check for regressions.** Before finalizing, run the full test suite:
           ./tools/run-tests.sh --test visual_comparison
           This produces compact output showing only regressions and improvements.

        7. **Accept baselines** for all changed scores:
           ./tools/target/debug/accept-baselines

        ## Token efficiency — IMPORTANT
        - ALWAYS use `./tools/run-tests.sh` instead of raw `cargo test` — it produces ~2 lines instead of ~666
        - Use `./tools/run-tests.sh --verbose` ONLY if you need full output for debugging
        - When reading source files, read only the specific line ranges you need, not entire files
        - When inspecting DOCX XML, dump only the specific internal file you need, not all of them
        - Do NOT re-read files you've already read unless you've made changes to them
        - Keep bash command output short — use head/tail/grep to filter

        ## Finalization — MANDATORY

        **Step 1: Update the case history log.** Append a summary of this run to the persistent history file at `{history_path}`. This file persists across runs so future agents learn from your work. Write in plain text, appending to whatever is already there. Include:
        - Date and starting/ending scores
        - What you investigated and found (DOCX features, root causes)
        - What you tried and whether it worked or not
        - What you changed and why
        - What remains to be done
        - Any dead ends future agents should avoid

        **Step 2: Write the outcome file.**
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
        - **"summary"** MUST describe what you actually did:
          - If you changed code: list each change briefly (e.g. "Implemented fldSimple PAGE/NUMPAGES fields in header_footer.rs; added WMF fallback rendering in images.rs")
          - If you made NO code changes: explicitly state "NO CODE CHANGES" and explain why (e.g. "NO CODE CHANGES — all remaining differences are systemic text spacing drift")
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
        - Run tests (compact): ./tools/run-tests.sh --case {case_name} --test visual_comparison
        - Run tests (verbose): ./tools/run-tests.sh --case {case_name} --test visual_comparison --verbose
        - Run full suite: ./tools/run-tests.sh
        - Inspect DOCX: ./tools/target/debug/docx-inspect tests/fixtures/{case_path}/input.docx [path]
        - Analysis: ./tools/target/debug/analyze-fixtures --grep "pattern"
    """)


def build_discovery_prompt(case_name: str, case_path: str, progress_file: Path, logs_dir: Path) -> str:
    """Prompt for newly discovered fixtures — emphasizes feature analysis over score chasing."""
    history = get_case_history(case_name)
    history_path = history_file_for(case_name)

    history_section = ""
    if history:
        history_section = (
            "\n        ## History from previous runs\n"
            "        The following is a log from previous agent runs on this case. Read it carefully —"
            " it documents what was tried, what worked, what failed, and what was learned."
            " Do NOT repeat failed approaches.\n\n"
            f"        ```\n{history}\n        ```\n"
        )

    return dedent(f"""\
        You are working on the docxside-pdf project — a Rust library that converts DOCX files to PDF.
        This is a **newly discovered** test fixture: **{case_name}** (path: tests/fixtures/{case_path}/).
        It has never been tested before. Your job is to analyze it thoroughly and fix any rendering issues that are caused by
        unhandled features, missing implementations, or bugs in our code.
        {history_section}
        ## Your goal
        1. Run the test and establish a baseline score
        2. Thoroughly analyze what DOCX features this document uses
        3. Identify which features we handle poorly or not at all
        4. **Fix everything you can.** Implement missing features, fix bugs, handle edge cases. Nothing is "out of scope" except the systemic issues listed below.
        5. Only terminate early if the score gap is ENTIRELY caused by systemic text rendering issues (spacing, kerning, line height, drift) — meaning there is literally nothing else wrong.

        ## WHAT TO FIX — everything except systemic text rendering
        You must attempt to fix ANY issue that is not in the narrow systemic exclusion list below. This includes but is not limited to:
        - Missing borders, backgrounds, shading, or fills
        - Missing or broken images, shapes, or drawings (including WMF/EMF — implement basic support if needed)
        - Missing text effects (bold, italic, underline, strikethrough, smallcaps, allcaps, etc.)
        - Entirely missing content blocks (paragraphs, table rows, sections)
        - Wrong colors or wrong fonts being selected
        - Tables with missing cells, wrong column widths, or missing gridlines
        - Headers/footers not rendered or rendered incorrectly
        - Field codes (fldSimple, SECTIONPAGES, NUMPAGES, DOCPROPERTY, PAGE, etc.) — implement them
        - List numbering or bullet issues
        - Any DOCX feature that we simply don't handle at all — implement it
        - Anything that looks like a bug (wrong rendering) vs. imprecision (slightly off rendering)
        - Even "small pixel impact" issues — fix them. Every improvement counts.

        **Do NOT dismiss issues as "out of scope" or "too small to matter".** If you can see a difference in the diff that isn't systemic spacing, you should try to fix it. Implement new features if needed. The whole point of this project is to get closer to Word's output.

        ## EARLY TERMINATION — only for truly systemic-only cases
        You may ONLY terminate early if, after investigation, the diff images show that our output has the **right structure** and the ONLY remaining differences are ALL of these systemic types:
        - Word/character spacing or kerning precision
        - Line height or line spacing drift
        - Vertical or horizontal positional drift/accumulation
        - Font metric differences (slightly different glyph widths)
        - Page break position differences caused by accumulated spacing drift

        If there is ANYTHING else wrong — even one missing field code, one unsupported image format, one missing border — you must attempt to fix it before terminating.

        ## Progress logging
        After each significant action, append a timestamped entry to your progress file:
          echo "$(date '+%H:%M:%S') — <what you did and what happened>" >> "{progress_file}"

        ## Workflow
        1. **Run the initial test** to establish a baseline:
           ./tools/run-tests.sh --case {case_name} --test visual_comparison
           Log the starting scores.

        2. **Inspect the fixture thoroughly.** Use `./tools/target/debug/docx-inspect tests/fixtures/{case_path}/input.docx` to list ZIP entries.
           Then examine key XML files:
           - `word/document.xml` — what elements and features are used?
           - `word/styles.xml` — any unusual styles?
           - Check for: tables, images, charts, shapes, textboxes, headers/footers, numbering, etc.

        3. **Analyze the diff images.** Look at tests/output/{case_path}/diff/ — blue = reference only, red = generated only.
           Be thorough — look at EVERY page. Categorize what you see:
           - **Fixable issues** (missing elements, wrong rendering, unsupported features, missing field codes, unsupported image formats, wrong colors, missing effects) → FIX THESE
           - **Spacing/drift issues** (content present and correct but shifted by a few pixels) → these are systemic, skip them
           Write your analysis to the progress file. When in doubt, it's fixable.

        4. **If ALL issues are truly systemic spacing/drift with zero fixable issues:** Skip to Finalization. This should be rare.

        5. **Fix everything you can.** Implement missing features, fix bugs, add support for new formats. After each change, run:
           ./tools/run-tests.sh --case {case_name} --test visual_comparison

        6. **Check for regressions** before finalizing:
           ./tools/run-tests.sh --test visual_comparison

        7. **Accept baselines** for all changed scores:
           ./tools/target/debug/accept-baselines

        ## Token efficiency — IMPORTANT
        - ALWAYS use `./tools/run-tests.sh` instead of raw `cargo test`
        - Use `./tools/run-tests.sh --verbose` ONLY if you need full output for debugging
        - Read only specific line ranges, not entire files
        - When inspecting DOCX XML, dump only the specific internal file you need
        - Keep bash command output short — use head/tail/grep to filter

        ## Finalization — MANDATORY

        **Step 1: Update the case history log.** Append a summary to `{history_path}`. Include:
        - Date and starting/ending scores
        - Full feature analysis: what DOCX features this document uses
        - What you investigated and found
        - What you tried and whether it worked
        - What remains to be done
        - Classification: "systemic-only" if no actionable issues, or list of specific feature gaps

        **Step 2: Write the outcome file.**
        Write this file as your very last action to: `{logs_dir / (case_name + ".outcome.json")}`

        ```json
        {{
          "case": "{case_name}",
          "case_path": "{case_path}",
          "improved": true,
          "target_before": {{"jaccard": 0.0, "ssim": 0.0}},
          "target_after": {{"jaccard": 0.0, "ssim": 0.0}},
          "regressions": [],
          "discovery_analysis": {{
            "features_found": ["list", "of", "DOCX", "features", "used"],
            "unhandled_features": ["features", "we", "dont", "support"],
            "systemic_only": false,
            "classification": "actionable|systemic-only|mixed"
          }},
          "summary": "Description of what you found and changed"
        }}
        ```

        Rules for the outcome:
        - **"improved": true** only if Jaccard OR SSIM increased by at least 0.5% (0.005)
        - **"regressions"** lists ANY other case where Jaccard or SSIM dropped by more than 2% (0.02)
        - **"discovery_analysis"** is specific to discovery mode — document what features the fixture uses
        - If `systemic_only` is true, set `improved` to false
        - **"summary"** MUST describe what you actually did:
          - If you changed code: list each change briefly (e.g. "Implemented fldSimple PAGE/NUMPAGES fields in header_footer.rs; added WMF fallback rendering in images.rs")
          - If you made NO code changes: explicitly state "NO CODE CHANGES" and explain why (e.g. "NO CODE CHANGES — all remaining differences are systemic text spacing drift")
        - **Commit your changes** (code + baselines.json) before writing the outcome file. End every commit message with:
          Automated-by: run-agents.py ({case_name})

        ## Important rules
        - Do NOT modify test fixtures or reference PDFs
        - Do NOT modify the test harness or scoring thresholds
        - Keep changes focused but DO implement new features when needed — adding a WMF renderer, implementing fldSimple fields, handling a new XML element are all in scope
        - If the case fails due to missing fonts, log that and move on to other issues you CAN fix
        - The DOCX spec can be queried via the local RAG tool (mcp__local-rag__query_documents)
        - **Nothing is "out of scope" except systemic text spacing.** If a feature is missing, implement it. If a format is unsupported, add support. That is the job.

        ## Key environment
        - Run tests (compact): ./tools/run-tests.sh --case {case_name} --test visual_comparison
        - Run tests (verbose): ./tools/run-tests.sh --case {case_name} --test visual_comparison --verbose
        - Run full suite: ./tools/run-tests.sh
        - Inspect DOCX: ./tools/target/debug/docx-inspect tests/fixtures/{case_path}/input.docx [path]
        - Analysis: ./tools/target/debug/analyze-fixtures --grep "pattern"
    """)


# ── Agent launcher ───────────────────────────────────────────────────────────


def meta_file_for(logs_dir: Path, case_name: str) -> Path:
    return logs_dir / f"{case_name}.meta.json"


def save_agent_meta(logs_dir: Path, case_name: str, case_path: str, discovery: bool) -> None:
    meta_file_for(logs_dir, case_name).write_text(json.dumps({
        "case": case_name,
        "case_path": case_path,
        "discovery": discovery,
    }))


def load_agent_meta(logs_dir: Path, case_name: str) -> dict | None:
    path = meta_file_for(logs_dir, case_name)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text())
    except Exception:
        return None


def build_resume_prompt(case_name: str, case_path: str, progress_file: Path, logs_dir: Path, discovery: bool) -> str:
    """Wrap the original prompt with a resume preamble pointing at prior work."""
    base = build_discovery_prompt(case_name, case_path, progress_file, logs_dir) if discovery \
        else build_prompt(case_name, case_path, progress_file, logs_dir)

    preamble = dedent(f"""\
        ## RESUMING INTERRUPTED WORK
        You were previously launched on this case but the run did not finish (no outcome file was written).
        Your worktree is intact and your prior progress log is at `{progress_file}`. Before doing anything else:

        1. Read the progress log to see what you already investigated, tried, and concluded.
        2. Run `git status` and `git diff main...HEAD` in the current directory to see uncommitted/committed work.
        3. Read the case history file (path mentioned below in the original prompt) for context from earlier runs.
        4. Continue from where you left off — do NOT redo investigation that's already in the log.
        5. The end goal is unchanged: improve scores for this case and write the outcome file as your last action.

        Append a fresh entry to the progress log noting that you've resumed and summarizing what you found
        in the existing log/diff before continuing.

        --- ORIGINAL TASK BRIEF (still applies) ---

    """)
    return preamble + base


def launch_agent(
    case_name: str,
    case_path: str,
    logs_dir: Path,
    model: str,
    max_turns: int | None,
    permission_mode: str,
    discovery: bool = False,
    resume: bool = False,
) -> dict:
    """Run a single agent in a worktree. Called in a subprocess via ProcessPoolExecutor."""
    progress_file = logs_dir / f"{case_name}.log"
    outcome_file = logs_dir / f"{case_name}.outcome.json"

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if resume:
        with open(progress_file, "a") as f:
            f.write(f"\n{timestamp} — Agent RESUMED for {case_name} ({case_path})\n")
    else:
        progress_file.write_text(f"{timestamp} — Agent started for {case_name} ({case_path})\n")

    try:
        wt_path = setup_worktree(case_name, resume=resume)
    except Exception as e:
        with open(progress_file, "a") as f:
            f.write(f"{timestamp} — Failed to setup worktree: {e}\n")
        return {"case": case_name, "error": f"worktree setup failed: {e}"}

    if not resume:
        save_agent_meta(logs_dir, case_name, case_path, discovery)

    if resume:
        prompt = build_resume_prompt(case_name, case_path, progress_file, logs_dir, discovery)
    elif discovery:
        prompt = build_discovery_prompt(case_name, case_path, progress_file, logs_dir)
    else:
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


def _extract_dict(data: dict, *keys) -> dict:
    """Try multiple key names, return the first that's a dict."""
    for k in keys:
        v = data.get(k)
        if isinstance(v, dict):
            return v
    return {}


def parse_outcome(outcome_file: str) -> dict | None:
    path = Path(outcome_file)
    if not path.exists():
        return None
    try:
        with open(path) as f:
            data = json.load(f)

        # Agents use varying field names — normalize them
        reg = data.get("regressions", [])

        # improved: bool field, or infer from "status" == "improved"
        improved = data.get("improved")
        if improved is None:
            improved = data.get("status") == "improved"

        # summary: try several field names
        summary = (
            data.get("summary")
            or data.get("changes_summary")
            or data.get("description")
            or "No summary"
        )

        # before/after scores: agents use target_before, before, score_before, etc.
        before = _extract_dict(data, "target_before", "before", "score_before")
        after = _extract_dict(data, "target_after", "after", "score_after")

        # Discovery analysis (optional, only in discover mode)
        discovery = data.get("discovery_analysis")

        return {
            "improved": bool(improved),
            "regressions": reg if isinstance(reg, list) else [],
            "summary": summary,
            "target_before": before,
            "target_after": after,
            "discovery_analysis": discovery,
        }
    except Exception as e:
        return {"error": str(e)}


def resolve_merge_conflict(case_name: str, branch: str, permission_mode: str, logs_dir: Path) -> bool:
    """Launch a Claude agent to resolve merge conflicts in REPO_ROOT. Returns True on success."""
    # Find conflicted files
    result = subprocess.run(
        ["git", "-C", str(REPO_ROOT), "diff", "--name-only", "--diff-filter=U"],
        capture_output=True, text=True, check=False,
    )
    conflicted = result.stdout.strip()
    if not conflicted:
        return False

    log_file = logs_dir / f"{case_name}.merge-resolve.log"

    prompt = dedent(f"""\
        You are resolving a git merge conflict in the docxside-pdf project.
        Branch `{branch}` is being merged into `main`. The merge has already been started but has conflicts.

        Conflicted files:
        {conflicted}

        ## Your task
        1. Read each conflicted file and understand both sides of the conflict.
        2. The branch side ({branch}) contains work by an agent that improved rendering for {case_name}.
           The main side contains changes from other agents that were merged first.
        3. Resolve each conflict by keeping BOTH sides' changes — they are almost certainly in different
           parts of the code (different functions, different match arms, different modules). If both sides
           changed the same lines, use your judgment to combine them correctly.
        4. For `baselines.json` or `visual_hashes.json` conflicts: keep ALL score/hash entries from both sides.
        5. After resolving, stage the files with `git add` and complete the merge with `git commit --no-edit`.
        6. Verify the result compiles: `cargo check`
        7. If cargo check fails, fix the compilation errors and amend the merge commit.

        ## Important
        - Do NOT abort the merge.
        - Do NOT drop either side's changes — combine them.
        - If you truly cannot resolve a conflict, leave the file conflicted and exit — the orchestrator will abort.
    """)

    cmd = [
        "claude", "-p",
        "--model", "sonnet",
        "--max-turns", "20",
        "--permission-mode", permission_mode,
        prompt,
    ]

    print(f"  [{case_name}] Launching merge conflict resolver...")
    with open(log_file, "w") as log:
        result = subprocess.run(
            cmd,
            cwd=str(REPO_ROOT),
            stdout=log,
            stderr=subprocess.STDOUT,
        )

    # Check if conflicts are resolved (no more unmerged files)
    check = subprocess.run(
        ["git", "-C", str(REPO_ROOT), "diff", "--name-only", "--diff-filter=U"],
        capture_output=True, text=True, check=False,
    )
    if check.stdout.strip():
        print(f"  [{case_name}] Merge resolver failed — conflicts remain")
        return False

    print(f"  [{case_name}] Merge conflicts resolved")
    return True


def try_merge(case_name: str, wt_path: str, permission_mode: str, logs_dir: Path) -> tuple[bool, str]:
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
        # Merge conflict — try to resolve automatically
        if resolve_merge_conflict(case_name, branch, permission_mode, logs_dir):
            pass  # conflict resolved, fall through to cleanup
        else:
            # Resolution failed — abort
            subprocess.run(
                ["git", "-C", str(REPO_ROOT), "merge", "--abort"],
                capture_output=True, check=False,
            )
            return False, f"Merge conflict (auto-resolve failed):\n{result.stdout}\n{result.stderr}"

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


def process_results(
    results: list[dict],
    logs_dir: Path,
    discovery_cases: set[str] | None = None,
    permission_mode: str = "auto",
):
    """Process agent results: auto-merge or flag for review."""
    if discovery_cases is None:
        discovery_cases = set()
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

        discovery = outcome.get("discovery_analysis")
        if discovery and isinstance(discovery, dict):
            classification = discovery.get("classification", "?")
            print(f"  Discovery: {classification}")
            unhandled = discovery.get("unhandled_features", [])
            if unhandled:
                print(f"  Unhandled features: {', '.join(unhandled)}")

        if outcome["improved"] and len(regs) == 0:
            print(f"  Improved with no regressions — auto-merging into main")
            ok, msg = try_merge(case_name, wt_path, permission_mode, logs_dir)
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
            if case_name in discovery_cases:
                print(f"  (fixture already committed to main — only code changes need review)")

        else:
            if case_name in discovery_cases:
                print(f"  No meaningful code improvement (fixture already committed to main)")
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


def find_resumable_globally(limit: int) -> list[tuple[Path, str, dict]]:
    """Resumable cases across all logs/agents-* dirs.

    For each unique case, picks the most recent logs dir that has a resumable record.
    Returns up to `limit` entries sorted by worktree mtime descending.
    """
    parent = REPO_ROOT / "logs"
    if not parent.is_dir():
        return []
    all_dirs = sorted((p for p in parent.glob("agents-*") if p.is_dir()), reverse=True)
    seen: set[str] = set()
    found: list[tuple[float, Path, str, dict]] = []
    for logs_dir in all_dirs:
        for case_name, meta in find_resumable_cases(logs_dir):
            if case_name in seen:
                continue
            seen.add(case_name)
            wt_path = REPO_ROOT / ".worktrees" / f"agent-{case_name}"
            mtime = wt_path.stat().st_mtime if wt_path.is_dir() else 0
            found.append((mtime, logs_dir, case_name, meta))
    found.sort(key=lambda x: x[0], reverse=True)
    return [(d, n, m) for _, d, n, m in found[:limit]]


def run_resume(args) -> None:
    """Re-launch agents for cases that have a worktree but no outcome file.

    --resume        → all resumable in the latest logs/agents-* dir
    --resume PATH   → all resumable in the given logs dir
    --resume N      → N most recent resumable worktrees across all logs dirs
    """
    triples: list[tuple[Path, str, dict]]
    raw = args.resume

    if raw and raw.isdigit():
        limit = int(raw)
        if limit <= 0:
            print("Error: --resume N must be a positive integer", file=sys.stderr)
            sys.exit(1)
        triples = find_resumable_globally(limit)
        if not triples:
            print("Nothing to resume — no worktrees with unfinished outcomes found.")
            return
        print(f"Resuming {len(triples)} most recent worktree(s) across all logs dirs.")
    elif raw:
        logs_dir = Path(raw).expanduser().resolve()
        if not logs_dir.is_dir():
            print(f"Error: --resume path is not a directory: {logs_dir}", file=sys.stderr)
            sys.exit(1)
        triples = [(logs_dir, n, m) for n, m in find_resumable_cases(logs_dir)]
    else:
        latest = find_latest_logs_dir()
        if latest is None:
            print("Error: no logs/agents-* directories found", file=sys.stderr)
            sys.exit(1)
        print(f"Resuming from latest logs dir: {latest}")
        triples = [(latest, n, m) for n, m in find_resumable_cases(latest)]

    if args.cases:
        wanted = set(args.cases)
        triples = [(d, n, m) for d, n, m in triples if n in wanted]
        missing = wanted - {n for _, n, _ in triples}
        if missing:
            print(f"Warning: not resumable (no meta/worktree, or outcome already exists): {', '.join(sorted(missing))}",
                  file=sys.stderr)

    if not triples:
        print("Nothing to resume — no matching cases.")
        return

    print(f"\nResuming {len(triples)} agent(s):")
    for logs_dir, name, meta in triples:
        mode = "discovery" if meta.get("discovery") else "regular"
        print(f"  - {name} ({mode}) — logs: {logs_dir}")
    print(f"\nMonitor progress:")
    for logs_dir in sorted({d for d, _, _ in triples}):
        print(f"  tail -f {logs_dir}/*.log")
    print()

    discovery_cases = {n for _, n, m in triples if m.get("discovery")}

    results: list[dict] = []
    if args.sequential:
        print("Sequential mode: resuming one case at a time\n")
        for logs_dir, name, meta in triples:
            try:
                r = launch_agent(
                    name,
                    meta["case_path"],
                    logs_dir,
                    args.model,
                    args.max_turns,
                    args.permission,
                    discovery=name in discovery_cases,
                    resume=True,
                )
            except Exception as e:
                print(f"  [{name}] Exception: {e}")
                r = {"case": name, "error": str(e)}
            r["_logs_dir"] = str(logs_dir)
            results.append(r)
    else:
        with ProcessPoolExecutor(max_workers=len(triples)) as pool:
            futures = {
                pool.submit(
                    launch_agent,
                    name,
                    meta["case_path"],
                    logs_dir,
                    args.model,
                    args.max_turns,
                    args.permission,
                    discovery=name in discovery_cases,
                    resume=True,
                ): (logs_dir, name)
                for logs_dir, name, meta in triples
            }
            for future in as_completed(futures):
                logs_dir, name = futures[future]
                try:
                    r = future.result()
                except Exception as e:
                    print(f"  [{name}] Exception: {e}")
                    r = {"case": name, "error": str(e)}
                r["_logs_dir"] = str(logs_dir)
                results.append(r)

    # Group by logs dir so process_results prints log paths correctly.
    order = {n: i for i, (_, n, _) in enumerate(triples)}
    by_dir: dict[str, list[dict]] = {}
    for r in results:
        by_dir.setdefault(r.pop("_logs_dir"), []).append(r)
    for logs_dir_str, group in by_dir.items():
        group.sort(key=lambda r: order.get(r["case"], 999))
        process_results(group, Path(logs_dir_str), discovery_cases=discovery_cases,
                        permission_mode=args.permission)


def find_latest_logs_dir() -> Path | None:
    parent = REPO_ROOT / "logs"
    if not parent.is_dir():
        return None
    candidates = sorted(p for p in parent.glob("agents-*") if p.is_dir())
    return candidates[-1] if candidates else None


def find_resumable_cases(logs_dir: Path) -> list[tuple[str, dict]]:
    """Cases in logs_dir that have a worktree but no outcome file.

    Prefers a .meta.json next to the .log; falls back to reconstructing meta
    from the .log filename (legacy runs that predate the meta file).
    """
    resumable: list[tuple[str, dict]] = []
    seen: set[str] = set()

    candidates: list[str] = []
    for log_path in sorted(logs_dir.glob("*.log")):
        # Skip helper logs like "<case>.merge-resolve.log"
        if log_path.name.count(".") != 1:
            continue
        candidates.append(log_path.stem)

    for case_name in candidates:
        if case_name in seen:
            continue
        seen.add(case_name)
        if (logs_dir / f"{case_name}.outcome.json").exists():
            continue
        wt_path = REPO_ROOT / ".worktrees" / f"agent-{case_name}"
        if not wt_path.is_dir():
            continue

        meta_path = meta_file_for(logs_dir, case_name)
        meta: dict | None = None
        if meta_path.exists():
            try:
                meta = json.loads(meta_path.read_text())
            except Exception:
                meta = None
        if meta is None:
            # Legacy run with no meta — reconstruct it.
            case_path = None
            for group_dir in sorted(FIXTURES_DIR.iterdir()):
                if group_dir.is_dir() and (group_dir / case_name).is_dir():
                    case_path = f"{group_dir.name}/{case_name}"
                    break
            if case_path is None:
                continue
            meta = {"case": case_name, "case_path": case_path, "discovery": False}

        resumable.append((case_name, meta))
    return resumable


def main():
    parser = argparse.ArgumentParser(description="Launch parallel Claude agents to improve test cases")
    parser.add_argument("cases", nargs="*", help="Case names to work on")
    parser.add_argument("--worst", type=int, default=0, help="Auto-select N worst-scoring cases from new/")
    parser.add_argument("--annotations", type=int, default=0, help="Auto-select N random cases with unfixed annotations")
    parser.add_argument("--discover", type=int, default=0,
                        help="Download N new DOCX files, name them, convert to PDF, and launch discovery agents")
    parser.add_argument("--resume", nargs="?", const="", default=None, metavar="N_OR_PATH",
                        help="Resume unfinished agents in their existing worktrees. "
                             "No arg: latest logs/agents-* dir. "
                             "Integer N: N most recent resumable worktrees across all logs dirs. "
                             "Path: a specific logs dir.")
    parser.add_argument("--model", default="opus", help="Claude model (default: opus)")
    parser.add_argument("--max-turns", type=int, default=None, help="Max conversation turns per agent")
    parser.add_argument("--permission", default="auto", help="Permission mode (default: auto)")
    parser.add_argument("--sequential", action="store_true",
                        help="Run agents one at a time instead of in parallel")
    args = parser.parse_args()

    if args.resume is not None:
        run_resume(args)
        return

    discovery_cases: set[str] = set()
    cases = list(args.cases)

    if args.discover > 0:
        print(f"\n{'=' * 60}")
        print(f"Discovery mode: downloading and preparing {args.discover} new fixtures...")
        print(f"{'=' * 60}\n")
        discovered = discover_new_cases(args.discover)
        cases.extend(discovered)
        discovery_cases.update(discovered)
        print(f"\nDiscovered {len(discovered)} new cases: {' '.join(discovered)}\n")

    if args.worst > 0:
        worst = pick_worst_cases(args.worst)
        cases.extend(worst)
        print(f"Auto-selected worst {args.worst} cases: {' '.join(worst)}")
    if args.annotations > 0:
        annotated = pick_annotated_cases(args.annotations)
        cases.extend(annotated)
        print(f"Auto-selected {len(annotated)} annotated cases: {' '.join(annotated)}")

    if not cases:
        parser.error("No cases specified. Use --worst N, --annotations N, --discover N, or pass case names.")

    # Resolve to group/case paths
    case_paths = {name: resolve_case_path(name) for name in cases}

    # Create logs directory
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    logs_dir = REPO_ROOT / "logs" / f"agents-{timestamp}"
    logs_dir.mkdir(parents=True, exist_ok=True)
    print(f"Logs: {logs_dir}")

    print(f"\nLaunching {len(cases)} agents (one per case)...")
    print(f"Cases: {' '.join(cases)}")
    if discovery_cases:
        print(f"Discovery mode: {' '.join(discovery_cases)}")
    print(f"\nMonitor progress:")
    print(f"  tail -f {logs_dir}/*.log\n")

    # Launch one agent per case, all in parallel (or sequentially if --sequential)
    results = []
    if args.sequential:
        print("Sequential mode: running one case at a time\n")
        for name in cases:
            try:
                results.append(launch_agent(
                    name,
                    case_paths[name],
                    logs_dir,
                    args.model,
                    args.max_turns,
                    args.permission,
                    discovery=name in discovery_cases,
                ))
            except Exception as e:
                print(f"  [{name}] Exception: {e}")
                results.append({"case": name, "error": str(e)})
    else:
        with ProcessPoolExecutor(max_workers=len(cases)) as pool:
            futures = {
                pool.submit(
                    launch_agent,
                    name,
                    case_paths[name],
                    logs_dir,
                    args.model,
                    args.max_turns,
                    args.permission,
                    discovery=name in discovery_cases,
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

    process_results(results, logs_dir, discovery_cases=discovery_cases, permission_mode=args.permission)


if __name__ == "__main__":
    main()
