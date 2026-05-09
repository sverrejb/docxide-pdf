# /// script
# requires-python = ">=3.10"
# dependencies = [
#   "msgraph-sdk>=1.0",
#   "azure-identity>=1.15",
#   "msal-extensions>=1.1",
# ]
# ///
"""Demo: DOCX → PDF via msgraph-sdk-python (official typed SDK).

Does the same upload → ?format=pdf → cleanup flow as a raw requests script, but uses:
- azure-identity.DeviceCodeCredential for sign-in (token persisted on disk so the
  second run is silent)
- GraphServiceClient for every Graph call — request paths are typed, responses
  come back as DriveItem objects instead of dicts

Usage:
    uv run tools/graph_sdk_demo.py <input.docx> <output.pdf>

Auth: first run prints a device code + URL. Sign in with a personal Outlook account
(the Sikt tenant blocks this app registration; see earlier IT ticket).
"""

import argparse
import asyncio
import sys
import uuid
from pathlib import Path

from azure.identity import DeviceCodeCredential, TokenCachePersistenceOptions
from msgraph import GraphServiceClient
from msgraph.generated.drives.item.items.item.content.content_request_builder import (
    ContentRequestBuilder,
)

CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
SCOPES = ["Files.ReadWrite"]


async def convert(docx: Path, pdf_out: Path) -> None:
    cred = DeviceCodeCredential(
        client_id=CLIENT_ID,
        cache_persistence_options=TokenCachePersistenceOptions(
            name="docxside-graph-sdk",
            allow_unencrypted_storage=True,
        ),
    )
    client = GraphServiceClient(credentials=cred, scopes=SCOPES)

    drive = await client.me.drive.get()
    drive_id = drive.id
    items = client.drives.by_drive_id(drive_id).items

    path = f"docxside-tmp/docxside-{uuid.uuid4().hex}-{docx.name}"
    path_ref = f"root:/{path}:"

    print(f"uploading to OneDrive:/{path}", file=sys.stderr)
    uploaded = await items.by_drive_item_id(path_ref).content.put(docx.read_bytes())
    item_id = uploaded.id

    try:
        print("requesting PDF conversion", file=sys.stderr)
        query_params = ContentRequestBuilder.ContentRequestBuilderGetQueryParameters(
            format="pdf"
        )
        config = ContentRequestBuilder.ContentRequestBuilderGetRequestConfiguration(
            query_parameters=query_params,
        )
        pdf_bytes = await items.by_drive_item_id(item_id).content.get(
            request_configuration=config,
        )
        if not pdf_bytes:
            raise RuntimeError("conversion returned no bytes")

        pdf_out.parent.mkdir(parents=True, exist_ok=True)
        pdf_out.write_bytes(pdf_bytes)
        print(f"wrote {pdf_out} ({len(pdf_bytes)} bytes)", file=sys.stderr)
    finally:
        await items.by_drive_item_id(item_id).delete()


def main() -> None:
    ap = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("docx", type=Path, help="input .docx")
    ap.add_argument("pdf", type=Path, help="output .pdf")
    args = ap.parse_args()

    if not args.docx.exists():
        sys.exit(f"not found: {args.docx}")

    asyncio.run(convert(args.docx, args.pdf))


if __name__ == "__main__":
    main()
