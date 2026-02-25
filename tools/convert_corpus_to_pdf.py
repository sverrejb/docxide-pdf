# /// script
# requires-python = ">=3.9"
# dependencies = []
# ///

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

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRAPED_DIR = PROJECT_ROOT / "tests" / "fixtures" / "scraped"
DEFAULT_DOWNLOAD_DIR = PROJECT_ROOT / "downloads"


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
    parser = argparse.ArgumentParser(description="Convert downloaded docx files to PDF and place in scraped fixtures.")
    parser.add_argument("batch", nargs="?", help="Folder with .docx files (default: downloads/)")
    args = parser.parse_args()

    input_dir = Path(args.batch) if args.batch else DEFAULT_DOWNLOAD_DIR
    if not input_dir.is_dir():
        raise SystemExit(f"Folder not found: {input_dir}")
    log.info("Converting from: %s", input_dir)

    SCRAPED_DIR.mkdir(parents=True, exist_ok=True)
    staging = Path.home() / "Documents" / f"_docx_convert_{uuid.uuid4().hex}"
    staging.mkdir()

    docx_files = list(input_dir.glob("*.docx"))
    log.info("Found %d files to convert", len(docx_files))

    try:
        for docx_file in docx_files:
            h = docx_file.stem
            log.info("Processing %s", h)

            doc_dir = SCRAPED_DIR / h
            doc_dir.mkdir(exist_ok=True)
            shutil.copy2(docx_file, doc_dir / "input.docx")

            tmp_docx = staging / docx_file.name
            tmp_pdf = staging / docx_file.with_suffix(".pdf").name
            shutil.copy2(docx_file, tmp_docx)
            subprocess.run(["xattr", "-d", "com.apple.quarantine", str(tmp_docx)], capture_output=True)

            try:
                convert_to_pdf(tmp_docx.resolve(), tmp_pdf.resolve())
                if tmp_pdf.exists():
                    shutil.move(str(tmp_pdf), doc_dir / "reference.pdf")
                    log.info("OK   %s", h)
                else:
                    log.warning("FAIL %s — PDF not created", h)
                    shutil.rmtree(doc_dir)
            except Exception as e:
                log.error("FAIL %s — %s", h, e)
                shutil.rmtree(doc_dir)
    finally:
        shutil.rmtree(staging)

    log.info("Done. Fixtures in %s", SCRAPED_DIR)


if __name__ == "__main__":
    main()
