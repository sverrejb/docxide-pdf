#!/usr/bin/env python3
"""Export .docx files to PDF with Microsoft Word for Mac, unattended.

    tools/word_export.py FILE_OR_DIR... [--out DIR] [--force] [--timeout S] [--keep-word] [--preset online|print]
    tools/word_export.py --check            # preflight only, converts nothing

Writes <stem>.pdf beside each input (or into --out). Exit code 1 if any file failed.

Why it needs no human at the keyboard:
- Inputs and outputs are staged inside Word's own sandbox container, so Word
  never shows its "Grant File Access" sheet (that sheet cannot be answered from
  a script without stealing focus, and a dismissed one grants nothing).
- Alerts are off, and a watchdog answers any dialog that appears anyway: Word's
  repair prompt is declined (a malformed file is a FAIL, not a wedge), error
  reporter and stray OK/Don't Save boxes are dismissed.
- Paths reach AppleScript as argv, never as source text, so quotes in names are fine.
- Each document has its own time budget. A timeout kills the osascript and
  recycles Word; three failures in a row also recycle Word.
- Only the document this script opened is ever closed. Word is quit at the end
  only if this script launched it.

One-time human steps (macOS, per terminal app): the first run asks whether the
terminal may control "Microsoft Word" (Automation) and "System Events"
(Accessibility, needed only by the watchdog).

Word's PDF "Optimize for" choice (the Save As radio button) lives in Office's
settings database, not in the AppleScript call. --preset (default online, the
tagged "Best for electronic distribution and accessibility" export of the
fixture references) is written there only when it differs, with Word closed.
"""
from __future__ import annotations

import argparse
import sqlite3
import shutil
import subprocess
import sys
import threading
import time
import uuid
import zipfile
from pathlib import Path

CONTAINER_TMP = Path.home() / "Library/Containers/com.microsoft.Word/Data/tmp"
WORD_APP = Path("/Applications/Microsoft Word.app")
REGISTRY_DIR = Path.home() / "Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB"
WORD_OPTIONS = ("Software", "Microsoft", "Office", "16.0", "Word", "Options")
# 1 = "Best for electronic distribution and accessibility" (Microsoft's online
# conversion service, tagged PDF), 0 = "Best for printing" (local, untagged).
ONLINE_PDF = "Use BCS Service to create tagged PDF"

# `display alerts` takes Word's VBA constants: 0 = wdAlertsNone.
EXPORT = """
on run argv
    set srcPath to item 1 of argv
    set dstPath to item 2 of argv
    set docName to item 3 of argv
    tell application "Microsoft Word"
        set display alerts to 0
        open POSIX file srcPath
        -- right after a (re)launch Word returns from open before the document is registered
        set waited to 0
        repeat until (exists document docName) or waited > 20
            delay 0.5
            set waited to waited + 0.5
        end repeat
        if not (exists document docName) then error "Word did not open " & docName
        set theDoc to document docName
        try
            -- "Show revisions in balloons" is a Word preference, not in the docx;
            -- the markup references want Word's inline All Markup view
            if (count of argv) > 3 then
                if item 4 of argv is "inline" then
                    set revisions mode of view of active window to in line revisions
                else
                    set revisions mode of view of active window to balloon revisions
                end if
            end if
            save as theDoc file name dstPath file format format PDF
            close theDoc saving no
        on error msg number n
            try
                close theDoc saving no
            end try
            error msg number n
        end try
    end tell
end run
"""

PING = 'tell application "Microsoft Word" to get name'
QUIT = 'tell application "Microsoft Word" to quit saving no'

# Buttons in the order they are tried. "No" first: that is the repair prompt
# ("Word found unreadable content… Do you want to recover?"), and recovering
# would silently change the document. Never "Yes"/"Select…": those repair or
# start a grant flow we do not want.
WATCHDOG = """
on run argv
    tell application "System Events"
        repeat with procName in {"Microsoft Word", "Microsoft Error Reporting"}
            if exists process procName then
                tell process procName
                    repeat with w in windows
                        repeat with label in {"No", "Don't Save", "Don't Send", "OK", "Cancel"}
                            if exists (button label of w) then
                                click (button label of w)
                                return label & " in " & procName
                            end if
                            if exists sheet 1 of w then
                                if exists (button label of sheet 1 of w) then
                                    click (button label of sheet 1 of w)
                                    return label & " in sheet of " & procName
                                end if
                            end if
                        end repeat
                    end repeat
                end tell
            end if
        end repeat
    end tell
    return ""
end run
"""


def osa(script: str, *args: str, timeout: float = 30) -> subprocess.CompletedProcess:
    return subprocess.run(["osascript", "-e", script, *args], capture_output=True, text=True, timeout=timeout)


def word_running() -> bool:
    # pgrep, not System Events: a hung System Events (AppleEvent timeout -1712)
    # must not stop the export, it only costs the dialog watchdog.
    return subprocess.run(["pgrep", "-xq", "Microsoft Word"]).returncode == 0


def registry_db() -> Path:
    dbs = sorted(REGISTRY_DIR.glob("MicrosoftRegistrationDB_*.reg"))
    if len(dbs) != 1:
        raise RuntimeError(f"expected one Office settings database in {REGISTRY_DIR}, found {len(dbs)}")
    return dbs[0]


def word_options_node(db: sqlite3.Connection) -> int:
    node = -1
    for name in WORD_OPTIONS:
        row = db.execute("SELECT node_id FROM HKEY_CURRENT_USER WHERE parent_id = ? AND name = ?",
                         (node, name)).fetchone()
        if row is None:
            raise RuntimeError("Word options key missing from the Office settings database: "
                               "save one PDF from Word by hand first")
        node = row[0]
    return node


def pdf_online() -> bool | None:
    """Word's current PDF preset: True online, False print, None never chosen."""
    with sqlite3.connect(f"file:{registry_db()}?mode=ro", uri=True) as db:
        row = db.execute("SELECT value FROM HKEY_CURRENT_USER_values WHERE node_id = ? AND name = ?",
                         (word_options_node(db), ONLINE_PDF)).fetchone()
    return None if row is None else bool(row[0])


def set_pdf_online(online: bool) -> None:
    with sqlite3.connect(registry_db()) as db:
        # type 4 = REG_DWORD, as Word writes it
        db.execute("INSERT OR REPLACE INTO HKEY_CURRENT_USER_values (node_id, name, type, value) "
                   "VALUES (?, ?, 4, ?)", (word_options_node(db), ONLINE_PDF, int(online)))


def preflight() -> list[str]:
    problems = []
    if not WORD_APP.exists():
        problems.append(f"{WORD_APP} not found")
    if not CONTAINER_TMP.is_dir():
        problems.append(f"{CONTAINER_TMP} missing: Word has never run on this account?")
    r = osa(PING, timeout=60)
    if r.returncode != 0:
        err = r.stderr.strip()
        if "-1743" in err:
            problems.append("Automation not granted: System Settings → Privacy & Security → Automation "
                            "→ allow this terminal to control Microsoft Word (or `tccutil reset AppleEvents`)")
        else:
            problems.append(f"osascript cannot reach Word: {err}")
    return problems


def warm(timeout: float = 90) -> None:
    """Launch Word in the background if needed and wait until it answers."""
    if not word_running():
        subprocess.run(["open", "-g", "-a", "Microsoft Word"], check=False)
    t0 = time.monotonic()
    while time.monotonic() - t0 < timeout:
        if osa(PING, timeout=20).returncode == 0:
            return
        time.sleep(1)
    raise RuntimeError("Word did not answer within %.0fs" % timeout)


def recycle() -> None:
    osa(QUIT, timeout=30)
    for _ in range(20):
        if not word_running():
            break
        time.sleep(0.5)
    else:
        subprocess.run(["pkill", "-9", "-x", "Microsoft Word"], check=False)
        time.sleep(2)
    warm()


class Watchdog(threading.Thread):
    def __init__(self, log):
        super().__init__(daemon=True)
        self.stop = threading.Event()
        self.log = log

    def run(self):
        seen = set()
        while not self.stop.is_set():
            try:
                r = osa(WATCHDOG, timeout=10)
                if r.stdout.strip():
                    self.log(f"  watchdog pressed {r.stdout.strip()}")
                err = r.stderr.strip()
                if err and err not in seen:  # once per distinct error, so a denied Accessibility grant is visible
                    seen.add(err)
                    self.log(f"  watchdog error: {err}")
            except subprocess.TimeoutExpired:
                self.log("  watchdog poll timed out")
            self.stop.wait(0.5)


def export_one(src: Path, dst: Path, stage: Path, timeout: float, revisions: str | None = None) -> None:
    # A broken zip makes Word ask to repair it, and that prompt needs a human
    # whenever the watchdog lacks Accessibility access.
    try:
        with zipfile.ZipFile(src) as z:
            bad = z.testzip()
    except zipfile.BadZipFile as e:
        raise RuntimeError(f"not a valid docx: {e}") from None
    if bad:
        raise RuntimeError(f"not a valid docx: corrupt part {bad}")
    # The document keeps its own name (FILENAME fields print it); a fresh
    # folder per export keeps the path unique.
    folder = stage / uuid.uuid4().hex[:6]
    folder.mkdir()
    s_docx, s_pdf = folder / src.name, folder / f"{src.stem}.pdf"
    shutil.copy2(src, s_docx)
    subprocess.run(["xattr", "-d", "com.apple.quarantine", str(s_docx)], capture_output=True)
    try:
        r = osa(EXPORT, str(s_docx), str(s_pdf), s_docx.name, *([revisions] if revisions else []), timeout=timeout)
        if r.returncode != 0:
            raise RuntimeError(r.stderr.strip() or f"osascript exit {r.returncode}")
        if not s_pdf.exists():
            raise RuntimeError("Word reported success but wrote no PDF")
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(s_pdf), dst)
    finally:
        shutil.rmtree(folder, ignore_errors=True)


def iter_docx(paths: list[Path]):
    for p in paths:
        if p.is_dir():
            yield from sorted(f for f in p.rglob("*.docx") if not f.name.startswith("~$"))
        elif p.suffix.lower() == ".docx" and not p.name.startswith("~$"):
            yield p
        else:
            print(f"skip {p}: not a .docx", file=sys.stderr)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("inputs", nargs="*", type=Path, help=".docx files or directories (recursive)")
    ap.add_argument("--out", type=Path, help="write PDFs here instead of beside the inputs")
    ap.add_argument("--force", action="store_true", help="overwrite existing PDFs")
    ap.add_argument("--timeout", type=float, default=120, help="seconds per document (default 120)")
    ap.add_argument("--keep-word", action="store_true", help="leave Word running afterwards")
    ap.add_argument("--check", action="store_true", help="preflight only (starts Word if it is not running)")
    ap.add_argument("--preset", choices=("online", "print"), default="online",
                    help="PDF \"Optimize for\": online = tagged, via Microsoft's service (default); print = local")
    ap.add_argument("--revisions", choices=("inline", "balloon"),
                    help="tracked changes: inline (All Markup in the text, as the corpus references) or balloon; "
                         "default = Word's current setting")
    a = ap.parse_args()

    problems = preflight()
    for p in problems:
        print(f"preflight: {p}", file=sys.stderr)
    if a.check:
        print("preflight OK" if not problems else "preflight FAILED")
        return 1 if problems else 0
    if problems:
        return 1
    if not a.inputs:
        ap.error("no inputs")

    jobs = []
    for src in iter_docx(a.inputs):
        dst = (a.out / f"{src.stem}.pdf") if a.out else src.with_suffix(".pdf")
        if dst.exists() and not a.force:
            print(f"skip {src.name}: {dst.name} exists (use --force)")
            continue
        jobs.append((src, dst))
    if not jobs:
        return 0

    online = a.preset == "online"
    if pdf_online() != online:
        # Word may only read the setting at launch; quitting it here could lose the user's documents
        if word_running():
            print(f"preset: Word's PDF preset is not '{a.preset}'; quit Word and run again", file=sys.stderr)
            return 1
        set_pdf_online(online)
        print(f"preset: Word's PDF preset set to '{a.preset}'")

    launched = not word_running()
    warm()
    stage = CONTAINER_TMP / f"docxide_{uuid.uuid4().hex[:8]}"
    stage.mkdir()
    dog = Watchdog(print)
    dog.start()
    failed, streak = 0, 0
    try:
        for src, dst in jobs:
            t0 = time.monotonic()
            try:
                export_one(src, dst, stage, a.timeout, a.revisions)
                streak = 0
                print(f"OK   {src.name} -> {dst} ({time.monotonic() - t0:.1f}s)")
            except subprocess.TimeoutExpired:
                failed += 1
                streak += 1
                print(f"FAIL {src.name}: timeout after {a.timeout:.0f}s, recycling Word")
                recycle()
                continue
            except Exception as e:  # noqa: BLE001 — every failure is one FAIL line, never a stopped batch
                failed += 1
                streak += 1
                print(f"FAIL {src.name}: {e}")
            if streak >= 3:
                print("  three failures in a row, recycling Word")
                recycle()
                streak = 0
    finally:
        dog.stop.set()
        shutil.rmtree(stage, ignore_errors=True)
        if launched and not a.keep_word:
            osa(QUIT, timeout=30)
    print(f"done: {len(jobs) - failed} ok, {failed} failed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
