#!/usr/bin/env python3
# /// script
# requires-python = ">=3.9"
# dependencies = []
# ///
"""Download random DOCX files, name them with Claude, and place in tests/fixtures/new/.

Usage:
    uv run tools/add_new_cases.py 10          # download, name, and convert 10 new cases
    uv run tools/add_new_cases.py 10 --skip-convert  # download and name only, skip Word PDF

Requires: curl, claude CLI, MS Word (for PDF conversion)
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import random
import shutil
import subprocess
import threading
import time
import uuid
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
NEW_DIR = PROJECT_ROOT / "tests" / "fixtures" / "new"
DOWNLOAD_DIR = PROJECT_ROOT / "downloads"
MANIFEST = PROJECT_ROOT.parent / "docx-corpus" / "manifest.txt"

WML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


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


# ── Claude naming ────────────────────────────────────────────────────────────


def name_files_with_claude(file_snippets: dict[str, str], existing_names: set[str]) -> dict[str, str]:
    """Call claude -p once to name all files. Returns {hash: name}."""
    existing_list = ", ".join(sorted(existing_names)) if existing_names else "(none)"

    entries = []
    for h, text in file_snippets.items():
        # Truncate text to keep prompt reasonable
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

    log.info("Asking Claude to name %d files...", len(file_snippets))
    result = subprocess.run(
        ["claude", "-p", "--model", "haiku", "--output-format", "json", prompt],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        log.error("Claude naming failed: %s", result.stderr)
        return {}

    try:
        response = json.loads(result.stdout)
        # Handle both direct JSON and wrapped {"result": "..."} format
        if "result" in response and isinstance(response["result"], str):
            # claude --output-format json wraps in {"result": "..."}
            inner = response["result"]
            # Find JSON object in the response
            start = inner.find("{")
            end = inner.rfind("}") + 1
            if start >= 0 and end > start:
                return json.loads(inner[start:end])
        elif all(isinstance(v, str) for v in response.values()):
            return response
    except (json.JSONDecodeError, AttributeError):
        # Try to find JSON in raw stdout
        try:
            text = result.stdout
            start = text.find("{")
            end = text.rfind("}") + 1
            if start >= 0 and end > start:
                return json.loads(text[start:end])
        except json.JSONDecodeError:
            pass

    log.error("Could not parse Claude response: %s", result.stdout[:500])
    return {}


# ── Word PDF conversion ─────────────────────────────────────────────────────

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


def dialog_watcher(stop_event: threading.Event) -> None:
    while not stop_event.is_set():
        subprocess.run(["osascript", "-e", DISMISS_SCRIPT], capture_output=True)
        time.sleep(0.5)


def convert_to_pdf(tmp_docx: Path, tmp_pdf: Path) -> None:
    script = f"""
        tell application "Microsoft Word"
            set display alerts to -2
            open POSIX file "{tmp_docx}"
            delay 2
            set theDoc to document 1
            save as theDoc file name "{tmp_pdf}" file format format PDF
            close theDoc saving no
            set display alerts to 0
        end tell
    """
    stop_event = threading.Event()
    watcher = threading.Thread(target=dialog_watcher, args=(stop_event,), daemon=True)
    watcher.start()
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    stop_event.set()
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip())


# ── Main ─────────────────────────────────────────────────────────────────────


def main() -> None:
    parser = argparse.ArgumentParser(description="Download and name new DOCX test cases")
    parser.add_argument("count", type=int, help="Number of cases to add")
    parser.add_argument("--skip-convert", action="store_true", help="Skip Word PDF conversion")
    parser.add_argument("--manifest", type=Path, default=MANIFEST)
    parser.add_argument("--min-size", type=int, default=10_000, help="Min DOCX size in bytes (default: 10KB)")
    parser.add_argument("--max-size", type=int, default=500_000, help="Max DOCX size in bytes (default: 500KB)")
    args = parser.parse_args()

    if not args.manifest.exists():
        raise SystemExit(f"Manifest not found: {args.manifest}")

    all_hashes = [h.strip() for h in args.manifest.read_text().splitlines() if h.strip()]

    # Collect content hashes of all existing input.docx files across all groups
    fixtures_dir = PROJECT_ROOT / "tests" / "fixtures"
    existing_content_hashes: set[str] = set()
    existing_names: set[str] = set()
    for group_dir in fixtures_dir.iterdir():
        if group_dir.is_dir():
            for case_dir in group_dir.iterdir():
                if case_dir.is_dir():
                    existing_names.add(case_dir.name)
                    docx = case_dir / "input.docx"
                    if docx.exists():
                        existing_content_hashes.add(hashlib.sha256(docx.read_bytes()).hexdigest())

    # Also track manifest hashes we've already downloaded
    existing_manifest_hashes: set[str] = set()
    for f in DOWNLOAD_DIR.glob("*.docx"):
        existing_manifest_hashes.add(f.stem)

    available = [h for h in all_hashes if h not in existing_manifest_hashes]
    random.shuffle(available)
    log.info("%d in manifest, %d existing fixtures (%d content hashes), %d available",
             len(all_hashes), len(existing_names), len(existing_content_hashes), len(available))

    DOWNLOAD_DIR.mkdir(exist_ok=True)
    NEW_DIR.mkdir(parents=True, exist_ok=True)

    # ── Phase 1: Download candidates ─────────────────────────────────────
    # Download more than needed since some will be filtered out
    candidates: list[tuple[str, Path]] = []  # (hash, path)
    target = args.count * 3  # over-fetch to account for filtering

    for h in available:
        if len(candidates) >= target:
            break

        dest = DOWNLOAD_DIR / f"{h}.docx"
        if not dest.exists():
            log.info("Downloading %s...", h[:16])
            result = subprocess.run(
                ["curl", "-sf", "-o", str(dest), f"https://docxcorp.us/documents/{h}.docx"],
                capture_output=True,
            )
            if result.returncode != 0:
                continue

        size = dest.stat().st_size
        if size < args.min_size or size > args.max_size:
            continue
        if not zipfile.is_zipfile(dest):
            continue

        # Check content hash against all existing fixtures
        content_hash = hashlib.sha256(dest.read_bytes()).hexdigest()
        if content_hash in existing_content_hashes:
            log.info("  Duplicate content (matches existing fixture), skipping")
            continue
        existing_content_hashes.add(content_hash)

        text = extract_text_snippet(dest)
        if len(text.strip()) < 30:
            continue

        candidates.append((h, dest))

    log.info("Downloaded %d valid candidates for %d slots", len(candidates), args.count)

    if len(candidates) < args.count:
        log.warning("Only found %d valid candidates (wanted %d)", len(candidates), args.count)

    # Trim to what we need
    candidates = candidates[:args.count]

    # ── Phase 2: Name all files with Claude ──────────────────────────────
    snippets = {}
    for h, path in candidates:
        snippets[h] = extract_text_snippet(path)

    names = name_files_with_claude(snippets, existing_names)
    if not names:
        raise SystemExit("Claude naming failed — no names returned")

    # Validate names
    used_names: set[str] = set(existing_names)
    hash_to_name: dict[str, str] = {}
    for h, _ in candidates:
        name = names.get(h)
        if not name or name in used_names:
            log.warning("No valid name for %s (got %r), skipping", h[:16], name)
            continue
        # Sanitize: ensure it's valid snake_case
        name = name.strip().lower().replace(" ", "_").replace("-", "_")
        name = "".join(c for c in name if c.isalnum() or c == "_")
        name = name.strip("_")
        if not name or name in used_names:
            log.warning("Sanitized name collision for %s: %r, skipping", h[:16], name)
            continue
        hash_to_name[h] = name
        used_names.add(name)

    log.info("Named %d files:", len(hash_to_name))
    for h, name in hash_to_name.items():
        log.info("  %s -> %s", h[:16], name)

    # ── Phase 3: Create fixtures and convert to PDF ──────────────────────
    staging = Path.home() / "Documents" / f"_docx_convert_{uuid.uuid4().hex}"
    if not args.skip_convert:
        staging.mkdir()

    added = 0
    try:
        for h, docx_path in candidates:
            name = hash_to_name.get(h)
            if not name:
                continue

            case_dir = NEW_DIR / name
            case_dir.mkdir(exist_ok=True)
            shutil.copy2(docx_path, case_dir / "input.docx")

            if not args.skip_convert:
                tmp_docx = staging / f"{name}.docx"
                tmp_pdf = staging / f"{name}.pdf"
                shutil.copy2(docx_path, tmp_docx)
                subprocess.run(["xattr", "-d", "com.apple.quarantine", str(tmp_docx)], capture_output=True)

                try:
                    convert_to_pdf(tmp_docx.resolve(), tmp_pdf.resolve())
                    if tmp_pdf.exists():
                        shutil.move(str(tmp_pdf), case_dir / "reference.pdf")
                        log.info("  %s — PDF OK", name)
                    else:
                        log.warning("  %s — PDF not created, removing", name)
                        shutil.rmtree(case_dir)
                        continue
                except Exception as e:
                    log.error("  %s — conversion failed: %s, removing", name, e)
                    shutil.rmtree(case_dir)
                    continue
            else:
                log.info("  %s — saved (no PDF)", name)

            added += 1
    finally:
        if staging.exists():
            shutil.rmtree(staging)

    log.info("Done. Added %d cases to %s", added, NEW_DIR)
    if args.skip_convert:
        log.info("PDF conversion was skipped. Convert manually with MS Word.")


if __name__ == "__main__":
    main()
