# /// script
# requires-python = ">=3.9"
# dependencies = []
# ///

"""Convert input.docx → reference.pdf for fixture directories that lack a reference.pdf.

Usage:
    uv run tools/convert_fixtures_to_pdf.py tests/fixtures/new
"""

import argparse
import logging
import shutil
import subprocess
import threading
import time
import uuid
from pathlib import Path

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)


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

    log.debug("Running osascript for %s", tmp_docx.name)
    result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)

    stop_event.set()
    if result.stderr:
        log.debug("osascript stderr: %s", result.stderr.strip())
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip())


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert input.docx to reference.pdf for fixture dirs missing a reference."
    )
    parser.add_argument("fixture_root", type=Path, help="Root folder containing fixture subdirectories")
    parser.add_argument("--force", action="store_true", help="Re-convert even if reference.pdf exists")
    args = parser.parse_args()

    if not args.fixture_root.is_dir():
        raise SystemExit(f"Not a directory: {args.fixture_root}")

    dirs = sorted(
        d for d in args.fixture_root.iterdir()
        if d.is_dir() and (d / "input.docx").exists()
    )

    to_convert = [
        d for d in dirs
        if args.force or not (d / "reference.pdf").exists()
    ]

    log.info("Found %d fixture dirs, %d need conversion", len(dirs), len(to_convert))
    if not to_convert:
        log.info("Nothing to do.")
        return

    staging = Path.home() / "Documents" / f"_docx_convert_{uuid.uuid4().hex}"
    staging.mkdir()

    try:
        for fixture_dir in to_convert:
            name = fixture_dir.name
            input_docx = fixture_dir / "input.docx"
            log.info("Converting %s", name)

            tmp_docx = staging / f"{name}.docx"
            tmp_pdf = staging / f"{name}.pdf"
            shutil.copy2(input_docx, tmp_docx)
            subprocess.run(["xattr", "-d", "com.apple.quarantine", str(tmp_docx)], capture_output=True)

            try:
                convert_to_pdf(tmp_docx.resolve(), tmp_pdf.resolve())
                if tmp_pdf.exists():
                    shutil.move(str(tmp_pdf), fixture_dir / "reference.pdf")
                    log.info("OK   %s", name)
                else:
                    log.warning("FAIL %s — PDF not created", name)
            except Exception as e:
                log.error("FAIL %s — %s", name, e)
    finally:
        shutil.rmtree(staging)

    log.info("Done.")


if __name__ == "__main__":
    main()
