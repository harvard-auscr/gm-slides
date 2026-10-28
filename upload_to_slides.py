"""Upload an already-generated deck to the Google Drive folder configured in .env.

    python upload_to_slides.py "output/GM #25 9.10---shoutouts.pptx"

generate_shoutouts.py does this automatically when .env enables it; this is the
retry path (network hiccup, expired sign-in) and the way to push an older deck.
Drive converts the PPTX into a Google Slides file on the way in. If a Slides
file with the deck's name is already in the folder you are asked before it is
overwritten in place (same link); --overwrite answers yes.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from shoutout_gen import drive

REPO_ROOT = Path(__file__).resolve().parent


def _main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("deck", type=Path, help="the .pptx to upload")
    parser.add_argument("--env-file", type=Path, default=REPO_ROOT / ".env", help="upload settings (see .env.example)")
    parser.add_argument(
        "--overwrite", action="store_true",
        help="replace a same-name Google Slides file in the folder without asking (default: ask; no answer keeps it)",
    )
    args = parser.parse_args(argv)

    try:
        cfg = drive.load_config(args.env_file)
    except drive.DriveConfigError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    if cfg is None:
        print(f"error: upload is not enabled; set {drive.KEY_FOLDER} in {args.env_file} (see .env.example)", file=sys.stderr)
        return 2
    on_existing = drive.always_overwrite if args.overwrite else drive.confirm_on_terminal
    try:
        uploaded = drive.upload_deck(args.deck, cfg, on_existing=on_existing)
    except FileNotFoundError:
        print(f"error: no such deck: {args.deck}", file=sys.stderr)
        return 2
    except drive.UploadDeclined as exc:
        print(f"not uploaded: {exc}: {exc.existing[0].link}\nrerun with --overwrite to replace it")
        return 0
    except Exception as exc:  # sign-in, network or Drive errors: the deck is untouched, just report
        print(f"upload failed: {exc}", file=sys.stderr)
        return 3
    verb = "overwrote" if uploaded.action == "overwritten" else "uploaded"
    print(f"{verb} {uploaded.name} as Google Slides: {uploaded.link}")
    return 0


if __name__ == "__main__":
    sys.exit(_main())
