# /// script
# requires-python = ">=3.9"
# dependencies = []
# ///

"""Convert input.docx → reference.pdf for fixture directories that lack a reference.pdf.

Usage:
    uv run tools/convert_fixtures_to_pdf.py tests/fixtures/new
    uv run tools/convert_fixtures_to_pdf.py --file tests/fixtures/cases/case1
    uv run tools/convert_fixtures_to_pdf.py --file path/to/some.docx
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


def _convert_one(staging: Path, name: str, input_docx: Path, output_pdf: Path) -> bool:
    tmp_docx = staging / f"{name}.docx"
    tmp_pdf = staging / f"{name}.pdf"
    shutil.copy2(input_docx, tmp_docx)
    subprocess.run(["xattr", "-d", "com.apple.quarantine", str(tmp_docx)], capture_output=True)

    try:
        convert_to_pdf(tmp_docx.resolve(), tmp_pdf.resolve())
    except Exception as e:
        log.error("FAIL %s — %s", name, e)
        return False

    if not tmp_pdf.exists():
        log.warning("FAIL %s — PDF not created", name)
        return False

    shutil.move(str(tmp_pdf), output_pdf)
    log.info("OK   %s", name)
    return True


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert input.docx to reference.pdf for fixture dirs missing a reference."
    )
    parser.add_argument("fixture_root", type=Path, nargs="?",
                        help="Root folder containing fixture subdirectories")
    parser.add_argument("--file", dest="single", type=Path,
                        help="Convert a single fixture dir (containing input.docx) or .docx file")
    parser.add_argument("--force", action="store_true", help="Re-convert even if reference.pdf exists")
    args = parser.parse_args()

    if bool(args.fixture_root) == bool(args.single):
        parser.error("Provide either a fixture_root positional or --file, not both.")

    staging = Path.home() / "Documents" / f"_docx_convert_{uuid.uuid4().hex}"
    staging.mkdir()

    try:
        if args.single:
            target = args.single
            if target.is_dir():
                input_docx = target / "input.docx"
                output_pdf = target / "reference.pdf"
                name = target.name
            elif target.is_file() and target.suffix.lower() == ".docx":
                input_docx = target
                output_pdf = target.with_name("reference.pdf")
                name = target.stem
            else:
                raise SystemExit(f"--file must be a fixture dir or .docx file: {target}")

            if not input_docx.exists():
                raise SystemExit(f"Missing input: {input_docx}")

            _convert_one(staging, name, input_docx, output_pdf)
            return

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

        for fixture_dir in to_convert:
            log.info("Converting %s", fixture_dir.name)
            _convert_one(staging, fixture_dir.name, fixture_dir / "input.docx", fixture_dir / "reference.pdf")
    finally:
        shutil.rmtree(staging)

    log.info("Done.")


if __name__ == "__main__":
    main()
