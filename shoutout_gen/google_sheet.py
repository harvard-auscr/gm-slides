"""Pull the shout-out responses straight from the Google Sheet behind the form.

Weekly use starts at the form's responses spreadsheet, so the generator accepts
that sheet's link in place of a downloaded file: the link names the spreadsheet
and, through ``gid``, the tab. The tab is fetched through the sheet's own CSV
export endpoint -- byte for byte what **File -> Download -> CSV** produces -- so
the reader sees exactly the shape it was calibrated on, and the download is
kept in ``input/`` like a manual export would be.

Why that endpoint and not an API (both were tried, 2026-09-23):
  * The Sheets API is a separate service that has to be enabled on the OAuth
    client's Cloud project (the club's enables only Drive), and it returns
    ragged value lists rather than the pinned CSV.
  * The Drive API's ``files.export`` ignores the tab: its CSV is always the
    first sheet, so a ``gid`` in the link would be dropped without a word.
The export endpoint accepts the OAuth token the upload already holds (the
``drive`` scope), so connecting the sheet costs no console work or new consent.
"""
from __future__ import annotations

import os
import re
import time
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path

from . import drive

KEY_SHEET = "SHOUTOUT_SHEET"  # the responses sheet's link in .env; makes the CLI's sheet argument optional

EXPORT_URL = "https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export"
# Sheet links come as /spreadsheets/d/<id>/edit... or /spreadsheets/u/<n>/d/<id>/...;
# the tab is the gid in the query or the fragment (a link from the address bar has both).
_SHEET_LINK = re.compile(r"docs\.google\.com/spreadsheets/(?:u/\d+/)?d/([A-Za-z0-9_-]+)")
_GID = re.compile(r"[?&#]gid=(\d+)")
# The endpoint has been seen to answer a valid request with a passing 400 (2026-09-23,
# while the owner was editing the sheet), so one failure earns one retry.
_ATTEMPTS = 2
_RETRY_DELAY_S = 2.0


class SheetLinkError(ValueError):
    """The text is not a Google Sheets link."""


class SheetDownloadError(RuntimeError):
    """Google did not hand back the tab as CSV (no access, wrong tab, expired sign-in, ...)."""


@dataclass(frozen=True)
class SheetRef:
    """One tab of one spreadsheet. ``gid`` None means the first tab, where a form writes its responses."""

    spreadsheet_id: str
    gid: str | None = None

    @property
    def link(self) -> str:
        base = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit"
        return f"{base}#gid={self.gid}" if self.gid is not None else base


def is_sheet_link(text: str) -> bool:
    return bool(_SHEET_LINK.search(str(text)))


def parse_sheet_link(text: str) -> SheetRef:
    """Spreadsheet ID and tab from a pasted Sheets link."""
    text = str(text).strip()
    match = _SHEET_LINK.search(text)
    if not match:
        raise SheetLinkError(f"not a Google Sheets link: {text!r}")
    gid = _GID.search(text)
    return SheetRef(match.group(1), gid.group(1) if gid else None)


def load_sheet_link(env_file: Path, environ: Mapping[str, str] = os.environ) -> str:
    """The sheet link from ``SHOUTOUT_SHEET`` (environment over .env), or "" when unset."""
    values = drive.read_env_file(Path(env_file))
    return (environ.get(KEY_SHEET) or values.get(KEY_SHEET, "")).strip()


def form_name_for(gm: int, date: str) -> str:
    """Where a pulled tab is kept, beside the deck's own name: ``GM #26 9.24---form.csv``."""
    return f"GM #{gm} {date}---form.csv"


def export_params(ref: SheetRef) -> dict[str, str]:
    """Query for the export endpoint: CSV, and the tab when the link named one."""
    params = {"format": "csv"}
    if ref.gid is not None:
        params["gid"] = ref.gid
    return params


def fetch_csv(ref: SheetRef, signin: drive.SignInConfig, session=None) -> bytes:
    """Download the tab as CSV bytes.

    ``session`` is injectable so tests can pass a stub; callers normally leave it
    ``None`` and get a signed-in ``requests`` session.
    """
    if session is None:
        session = _authorized_session(signin)
    url = EXPORT_URL.format(spreadsheet_id=ref.spreadsheet_id)
    for attempt in range(_ATTEMPTS):
        response = session.get(url, params=export_params(ref))
        content_type = (response.headers.get("content-type") or "").lower()
        if response.status_code == 200 and content_type.startswith("text/csv"):
            return response.content
        if attempt + 1 < _ATTEMPTS:
            time.sleep(_RETRY_DELAY_S)
    raise SheetDownloadError(_explain(response.status_code, content_type, ref))


def download(ref: SheetRef, signin: drive.SignInConfig, dest: Path, session=None) -> Path:
    """Fetch the tab and save it as ``dest`` (a .csv), creating the folder; returns ``dest``."""
    data = fetch_csv(ref, signin, session)
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(data)
    return dest


# --------------------------------------------------------------------------- #
# Private mechanics
# --------------------------------------------------------------------------- #
# Observed responses of the export endpoint (probed 2026-09-23): a wrong gid is a
# 400 HTML page, a missing/rejected token a 401 HTML page -- and, once, a 400 for
# a request that succeeded seconds later.
_HINTS = {
    400: "the tab (gid) in the link does not exist in that spreadsheet, or Google hiccupped (it was already retried once)",
    401: "the sign-in was not accepted; delete the cached token in .google/ and run again to sign in afresh",
    403: "the signed-in Google account cannot open that spreadsheet; share it with that account",
    404: "no spreadsheet with that ID",
}


def _explain(status: int, content_type: str, ref: SheetRef) -> str:
    hint = _HINTS.get(status, "Google returned something other than the tab's CSV")
    return f"could not download {ref.link}: HTTP {status}, {content_type or 'no content type'}: {hint}"


def _authorized_session(signin: drive.SignInConfig):
    try:
        from google.auth.transport.requests import AuthorizedSession
    except ImportError as exc:
        raise drive.DriveConfigError(
            "reading a Google Sheet needs the Google client libraries; run: pip install -r requirements.txt"
        ) from exc
    return AuthorizedSession(drive.sign_in(signin))
