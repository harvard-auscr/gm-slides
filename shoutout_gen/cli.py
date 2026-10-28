"""Command-line plumbing shared by the scripts: where the sheet comes from, the QOTW tally, and the Drive upload.

Each script parses its own options; these helpers take the parsed ``args`` (they
read ``sheet``, ``gm``, ``date``, ``input_dir``, ``env_file``, ``no_upload`` and
``overwrite``) and print errors the same way, returning an exit code on failure.
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from . import drive, google_sheet, term
from .qotw import Tally
from .sheet import parse_since

# The optional Google features read their settings from the repo's .env, and a
# pulled sheet lands in the repo's input/, regardless of the directory the
# command is run from.
REPO_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_ENV_FILE = REPO_ROOT / ".env"
DEFAULT_INPUT_DIR = REPO_ROOT / "input"


def since(value: str) -> datetime:
    """argparse type: the --since cutoff (see sheet.parse_since)."""
    try:
        return parse_since(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from exc


def add_source_args(parser: argparse.ArgumentParser) -> None:
    """The sheet, the GM it is for, the --since window and the Google settings."""
    parser.add_argument(
        "sheet", nargs="?",
        help="CSV/XLSX export of the form, or the link of the Google Sheet behind it "
             f"(default: {google_sheet.KEY_SHEET} in .env)",
    )
    parser.add_argument("--gm", type=int, required=True, help="general meeting number, e.g. 25")
    parser.add_argument("--date", required=True, help="date as written in deck names, e.g. 9.10")
    parser.add_argument(
        "--since", type=since, default=None,
        help="keep only responses submitted at or after this moment: m.d (e.g. 9.17), 'm.d HH:MM' or "
             "YYYY-MM-DD [HH:MM]; for a form that collects all year, pass when the previous GM's slide was "
             "made, e.g. '9.17 20:00' (a bare m.d means midnight, so that day's earlier rows come back)",
    )
    parser.add_argument("--out", type=Path, default=Path("output"), help="output folder")
    parser.add_argument(
        "--input-dir", type=Path, default=DEFAULT_INPUT_DIR,
        help="where a tab pulled from Google Sheets is saved as CSV before it is read (default: input/ in the repo)",
    )
    parser.add_argument("--env-file", type=Path, default=DEFAULT_ENV_FILE, help="Google settings (see .env.example)")


def add_upload_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--no-upload", action="store_true",
        help="skip the Drive upload even if .env enables it (a connected sheet is still read)",
    )
    parser.add_argument(
        "--overwrite", action="store_true",
        help="if a Google Slides file with the deck's name is already in the Drive folder, replace it without "
             "asking (default: ask on the terminal; no answer keeps the existing file)",
    )


def load_settings(args) -> tuple[drive.DriveConfig | None, str] | int:
    """(upload config or None, sheet source) from args and .env, or an exit code.

    Read before any work so a typo in .env fails fast instead of after the deck is made.
    """
    try:
        upload_cfg = None if args.no_upload else drive.load_config(args.env_file)
        source = args.sheet or google_sheet.load_sheet_link(args.env_file)
    except drive.DriveConfigError as exc:
        term.error(str(exc))
        return 2
    if not source:
        term.error(
            "nothing to read: pass the form export (CSV/XLSX) or the Google Sheet's link, "
            f"or set {google_sheet.KEY_SHEET} in {args.env_file}"
        )
        return 2
    return upload_cfg, source


def sheet_path(args, source: str) -> Path | int:
    """The file to read: ``source`` itself, or the tab it links to pulled into input/; or an exit code."""
    if not google_sheet.is_sheet_link(source):
        return Path(source)
    ref = google_sheet.parse_sheet_link(source)
    path = args.input_dir / google_sheet.form_name_for(args.gm, args.date)
    try:
        google_sheet.download(ref, drive.load_signin(args.env_file), path)
    except Exception as exc:  # sign-in, access or network: nothing has been generated yet, so just say why
        term.error(
            f"{exc}\nnothing was made. Download the tab yourself (File -> Download -> CSV) "
            "and pass the file instead"
        )
        return 2
    term.detail(f"pulled {ref.link} -> {path}")
    return path


def print_tally(result: Tally) -> None:
    """The question, then its answers one per line, most picked first."""
    print(f"{term.paint('QOTW:', 'bold', 'magenta')} {term.paint(result.question, 'bold')}")
    for name, n in result.counts:
        print(f"  {term.paint(f'{n:>3}', 'bold')}  {name}")


def upload(deck: Path, upload_cfg: drive.DriveConfig, overwrite: bool) -> int:
    """Upload ``deck`` as Google Slides and report; the exit code for the run."""
    on_existing = drive.always_overwrite if overwrite else drive.confirm_on_terminal
    try:
        uploaded = drive.upload_deck(deck, upload_cfg, on_existing=on_existing)
    except drive.UploadDeclined as exc:
        print(
            f"{term.paint('not uploaded:', 'bold', 'yellow')} {exc}: {term.path(exc.existing[0].link)}\n"
            f'the new deck is at {deck}; to replace the Slides file: python upload_to_slides.py --overwrite "{deck}"'
        )
        return 0
    except Exception as exc:  # sign-in, network or Drive errors: the deck is safe on disk, so report and let the caller retry
        print(
            f"{term.paint('upload failed:', 'bold', 'red', stream=sys.stderr)} {exc}\n"
            f'the deck is saved locally; retry with: python upload_to_slides.py "{deck}"',
            file=sys.stderr,
        )
        return 3
    verb = "overwrote the Google Slides file in Drive" if uploaded.action == "overwritten" else "uploaded to Google Drive as Google Slides"
    print(f"{term.paint(verb + ':', 'bold', 'green')} {term.path(uploaded.link)}")
    return 0
