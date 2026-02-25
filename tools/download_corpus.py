# /// script
# requires-python = ">=3.9"
# dependencies = []
# ///

import argparse
import logging
import random
import subprocess
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRAPED_DIR = PROJECT_ROOT / "tests" / "fixtures" / "scraped"
DEFAULT_MANIFEST = PROJECT_ROOT.parent / "docx-corpus" / "manifest.txt"


def main() -> None:
    parser = argparse.ArgumentParser(description="Download random docx files from the corpus.")
    parser.add_argument("count", type=int, help="Number of files to download")
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST, help="Path to manifest.txt")
    args = parser.parse_args()

    if not args.manifest.exists():
        raise SystemExit(f"Manifest not found: {args.manifest}")

    all_hashes = [h.strip() for h in args.manifest.read_text().splitlines() if h.strip()]
    existing = {p.name for p in SCRAPED_DIR.iterdir() if p.is_dir()} if SCRAPED_DIR.is_dir() else set()
    available = [h for h in all_hashes if h not in existing]

    log.info("%d total in manifest, %d already scraped, %d available", len(all_hashes), len(existing), len(available))

    if args.count > len(available):
        raise SystemExit(f"Requested {args.count} but only {len(available)} new hashes available")

    selected = random.sample(available, args.count)

    download_dir = PROJECT_ROOT / "downloads"
    download_dir.mkdir(exist_ok=True)
    log.info("Downloading %d files to %s/", args.count, download_dir)

    for h in selected:
        url = f"https://docxcorp.us/documents/{h}.docx"
        dest = download_dir / f"{h}.docx"
        log.info("  %s", h)
        subprocess.run(["curl", "-s", "-o", str(dest), url], check=True)

    log.info("Done.")


if __name__ == "__main__":
    main()
