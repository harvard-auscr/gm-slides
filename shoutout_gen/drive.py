"""Optional upload of a generated deck to Google Drive as a Google Slides file.

The club's GM decks live in Google Slides, so the useful end state of a
generated PPTX is a Slides file sitting in a shared Drive folder, ready for
File -> Import slides. This module does exactly that one upload (Drive converts
the PPTX on the way in) and nothing else.

The feature is opt-in via a ``.env`` file: ``load_config`` returns ``None``
unless ``DRIVE_FOLDER`` is set and ``DRIVE_UPLOAD`` is not switched off, and the
Google client libraries are imported only when an upload actually runs, so the
generator needs neither credentials nor network access by default.

The sign-in is shared: ``load_signin`` / ``sign_in`` give the same credentials to
``google_sheet``, which pulls the form's responses tab with them, so connecting
the sheet costs no second consent.

Auth choices (both deliberate):
  * A user OAuth flow, not a service account: service accounts have no storage
    quota on a personal My Drive, so their uploads into a shared folder fail.
    The one-time browser consent caches a refresh token next to the config.
  * The full ``drive`` scope, not ``drive.file``: the narrower scope only sees
    files the app itself created, so it cannot place an upload into a folder
    the user picked by pasting a link (the parent lookup 404s).

Re-running a week: a Google Slides file with the deck's name may already be in
the folder. ``upload_deck`` then asks its ``on_existing`` callback (the CLIs pass
a terminal prompt, or "yes" for ``--overwrite``) and either replaces that file's
content in place -- same file ID, same link -- or raises ``UploadDeclined`` and
leaves the folder untouched. It never silently creates a second copy.
"""
from __future__ import annotations

import os
import re
import sys
from collections.abc import Callable, Mapping
from dataclasses import dataclass
from pathlib import Path

# --------------------------------------------------------------------------- #
# The .env contract
# --------------------------------------------------------------------------- #
KEY_FOLDER = "DRIVE_FOLDER"  # folder link or bare ID; the feature is off without it
KEY_SWITCH = "DRIVE_UPLOAD"  # off/false/0/no keeps the folder configured but skips uploads
KEY_CLIENT_ID = "DRIVE_CLIENT_ID"  # inline OAuth "Desktop app" client ...
KEY_CLIENT_SECRET = "DRIVE_CLIENT_SECRET"
KEY_CLIENT_FILE = "DRIVE_CLIENT_FILE"  # ... or the client JSON downloaded from the console
KEY_TOKEN_FILE = "DRIVE_TOKEN_FILE"  # cache for the signed-in token

DEFAULT_CLIENT_FILE = ".google/credentials.json"
DEFAULT_TOKEN_FILE = ".google/token.json"
_OFF_VALUES = {"0", "off", "false", "no"}

SCOPES = ["https://www.googleapis.com/auth/drive"]  # also what the Sheets CSV export accepts (google_sheet)
GOOGLE_SLIDES_MIME = "application/vnd.google-apps.presentation"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

_ENV_LINE = re.compile(r"^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$")
# Drive folder URLs come as /drive/folders/<id> or /drive/u/<n>/folders/<id>, with or
# without ?usp=... -- the ID is the only part that matters.
_FOLDER_URL = re.compile(r"drive\.google\.com/.*?/folders/([A-Za-z0-9_-]+)")
_BARE_ID = re.compile(r"[A-Za-z0-9_-]{10,}")
_QUOTES = "\"'"


class DriveConfigError(ValueError):
    """The .env is present but unusable (syntax, bad folder link, no way to sign in)."""


@dataclass(frozen=True)
class SignInConfig:
    """How to obtain Google credentials: the cached token, and the OAuth client for a fresh consent.

    Paths are absolute (resolved against the .env). Shared by the upload and the
    Google Sheet pull, which run on the same token.
    """

    token_file: Path
    client_file: Path
    client_id: str = ""
    client_secret: str = ""


@dataclass(frozen=True)
class DriveConfig:
    """Everything one upload needs. Paths are absolute (resolved against the .env)."""

    folder_id: str
    token_file: Path
    client_file: Path
    client_id: str = ""
    client_secret: str = ""


@dataclass(frozen=True)
class UploadResult:
    file_id: str
    name: str
    link: str
    action: str = "created"  # or "overwritten": an existing same-name file was replaced in place


@dataclass(frozen=True)
class ExistingFile:
    """A Google Slides file already in the folder with the deck's name."""

    file_id: str
    name: str
    link: str
    modified: str  # RFC 3339 as Drive reports it, e.g. 2026-09-23T21:40:12.000Z


class UploadDeclined(Exception):
    """A same-name Slides file is in the folder and the caller chose not to overwrite it."""

    def __init__(self, existing: list[ExistingFile]):
        super().__init__(f'"{existing[0].name}" is already in the Drive folder and was left untouched')
        self.existing = existing


OnExisting = Callable[[list[ExistingFile]], bool]


def read_env_file(path: Path) -> dict[str, str]:
    """Parse ``KEY=VALUE`` lines; blanks, ``#`` comments, quotes and ``export`` allowed.

    A missing file is simply ``{}`` (the feature is off). A malformed line is an
    error rather than silently skipped, so a typo cannot quietly disable the upload.
    """
    path = Path(path)
    if not path.exists():
        return {}
    try:
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as exc:  # e.g. UTF-16, what Windows PowerShell's Out-File writes
        raise DriveConfigError(f"{path} is not UTF-8 text; save it as UTF-8") from exc
    values: dict[str, str] = {}
    for lineno, raw in enumerate(text.splitlines(), 1):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        match = _ENV_LINE.match(line)
        if not match:
            raise DriveConfigError(f"{path}:{lineno}: expected KEY=VALUE, got {raw!r}")
        key, value = match.group(1), match.group(2).strip()
        close = value.find(value[0], 1) if value and value[0] in _QUOTES else -1
        if close > 0:
            # Quoted: verbatim up to the closing quote (a '#' inside is not a comment); a comment may follow.
            after = value[close + 1:].strip()
            if after and not after.startswith("#"):
                raise DriveConfigError(f"{path}:{lineno}: unexpected text after the quoted value in {raw!r}")
            value = value[1:close]
        else:
            value = value.split(" #", 1)[0].rstrip()  # unquoted: allow a trailing comment
        values[key] = value
    return values


def extract_folder_id(link_or_id: str) -> str:
    """Accept a pasted Drive folder URL or the bare folder ID."""
    text = link_or_id.strip()
    match = _FOLDER_URL.search(text)
    if match:
        return match.group(1)
    if _BARE_ID.fullmatch(text):
        return text
    raise DriveConfigError(f"{KEY_FOLDER} is not a Drive folder link or ID: {link_or_id!r}")


def _settings(env_file: Path, environ: Mapping[str, str]):
    """A getter over the merged environment + .env values, and the folder relative paths resolve against.

    Real environment variables override the file so CI or a one-off shell can steer
    it without editing .env. Relative paths resolve against the .env's folder, so
    the result does not depend on the working directory.
    """
    env_file = Path(env_file)
    values = read_env_file(env_file)

    def get(key: str) -> str:
        return (environ.get(key) or values.get(key, "")).strip()

    return get, env_file.resolve().parent


def load_signin(env_file: Path, environ: Mapping[str, str] = os.environ) -> SignInConfig:
    """The sign-in settings, whether or not the upload is enabled.

    Always returns a config: without a .env it is just the default token and client
    paths. Whether a sign-in can actually happen is decided when it is attempted
    (``sign_in`` says what to set if it cannot).
    """
    get, base = _settings(env_file, environ)
    return SignInConfig(
        token_file=base / (get(KEY_TOKEN_FILE) or DEFAULT_TOKEN_FILE),
        client_file=base / (get(KEY_CLIENT_FILE) or DEFAULT_CLIENT_FILE),
        client_id=get(KEY_CLIENT_ID),
        client_secret=get(KEY_CLIENT_SECRET),
    )


def load_config(env_file: Path, environ: Mapping[str, str] = os.environ) -> DriveConfig | None:
    """Return the upload settings, or ``None`` when the feature is off.

    Off means: no .env, no ``DRIVE_FOLDER``, or ``DRIVE_UPLOAD`` set to off/false/0/no.
    """
    get, _ = _settings(env_file, environ)
    if get(KEY_SWITCH).lower() in _OFF_VALUES:
        return None
    folder = get(KEY_FOLDER)
    if not folder:
        return None
    signin = load_signin(env_file, environ)
    return DriveConfig(
        folder_id=extract_folder_id(folder),
        token_file=signin.token_file,
        client_file=signin.client_file,
        client_id=signin.client_id,
        client_secret=signin.client_secret,
    )


# --------------------------------------------------------------------------- #
# Sign-in and upload (Google libraries imported lazily, see module doc)
# --------------------------------------------------------------------------- #
def client_config(client_id: str, client_secret: str) -> dict:
    """The "installed app" client description the OAuth flow expects, from an inline ID + secret."""
    return {
        "installed": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": ["http://localhost"],
        }
    }


def upload_metadata(deck_path: Path, folder_id: str) -> dict:
    """Drive file resource for the deck: same name as the file, converted to Slides on upload."""
    return {"name": Path(deck_path).stem, "parents": [folder_id], "mimeType": GOOGLE_SLIDES_MIME}


def existing_query(name: str, folder_id: str) -> str:
    """Drive search for a live (not trashed) Google Slides file of that name in the folder."""
    escaped = name.replace("\\", "\\\\").replace("'", "\\'")
    return f"name = '{escaped}' and '{folder_id}' in parents and mimeType = '{GOOGLE_SLIDES_MIME}' and trashed = false"


def find_existing(service, name: str, folder_id: str) -> list[ExistingFile]:
    """Same-name Slides files already in the folder, newest first."""
    listed = (
        service.files()
        .list(
            q=existing_query(name, folder_id),
            fields="files(id,name,webViewLink,modifiedTime)",
            orderBy="modifiedTime desc",  # so the newest is in the page even with more than 10 copies
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    found = [
        ExistingFile(f["id"], f.get("name", name), f.get("webViewLink", ""), f.get("modifiedTime", ""))
        for f in listed.get("files", [])
    ]
    return sorted(found, key=lambda f: f.modified, reverse=True)


def confirm_on_terminal(existing: list[ExistingFile]) -> bool:
    """The interactive ``on_existing``: ask on the terminal; only y/yes overwrites.

    No terminal to answer on (piped or closed stdin) counts as no, with a hint
    about ``--overwrite`` so a script can opt in explicitly.
    """
    newest = existing[0]
    copies = f" ({len(existing)} copies; the newest one)" if len(existing) > 1 else ""
    when = f"{newest.modified[:16].replace('T', ' ')} UTC" if newest.modified else "date unknown"
    try:
        answer = input(f'"{newest.name}" is already in the Drive folder{copies}, last modified {when}. Overwrite it in place? [y/N] ')
    except EOFError:
        print("\nno answer (no terminal): the existing file is kept; pass --overwrite to replace it without asking", file=sys.stderr)
        return False
    return answer.strip().lower() in ("y", "yes")


def always_overwrite(existing: list[ExistingFile]) -> bool:
    """The ``on_existing`` behind ``--overwrite``."""
    return True


def _google_libs():
    """Import the Google client libraries, with an actionable error if they are missing."""
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
    except ImportError as exc:
        raise DriveConfigError(
            "Drive upload is configured but the Google client libraries are not installed; "
            "run: pip install -r requirements.txt"
        ) from exc
    return build, MediaFileUpload


def sign_in(cfg: SignInConfig | DriveConfig):
    """Cached token -> silent refresh -> one-time browser consent. Returns Google credentials.

    Takes anything with the ``SignInConfig`` fields, so the upload config works too.
    """
    from google.auth.exceptions import RefreshError
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    creds = None
    if cfg.token_file.exists():
        creds = Credentials.from_authorized_user_file(str(cfg.token_file), SCOPES)
    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            creds = None  # revoked or stale: fall through to a fresh consent
    if not (creds and creds.valid):
        if cfg.client_id and cfg.client_secret:
            flow = InstalledAppFlow.from_client_config(client_config(cfg.client_id, cfg.client_secret), SCOPES)
        elif cfg.client_file.exists():
            flow = InstalledAppFlow.from_client_secrets_file(str(cfg.client_file), SCOPES)
        else:
            raise DriveConfigError(
                f"no way to sign in: set {KEY_CLIENT_ID} + {KEY_CLIENT_SECRET} in .env, or point "
                f"{KEY_CLIENT_FILE} at the OAuth client JSON (looked for {cfg.client_file})"
            )
        creds = flow.run_local_server(port=0)
    cfg.token_file.parent.mkdir(parents=True, exist_ok=True)
    cfg.token_file.write_text(creds.to_json(), encoding="utf-8")
    return creds


def upload_deck(deck_path: Path, cfg: DriveConfig, service=None, on_existing: OnExisting | None = None) -> UploadResult:
    """Upload one PPTX into the configured folder as a Google Slides file.

    If a Slides file with the deck's name is already there, ``on_existing`` gets
    the matches (newest first) and returns True to overwrite the newest in place
    (same file ID and link) or False to leave the folder alone, in which case
    ``UploadDeclined`` is raised. With no ``on_existing`` an existing file is never
    touched -- the safe default for anything unattended.

    ``service`` is injectable so tests can pass a stub; callers normally leave it
    ``None`` and get a signed-in Drive v3 client.
    """
    deck_path = Path(deck_path)
    if not deck_path.is_file():
        raise FileNotFoundError(deck_path)
    build, MediaFileUpload = _google_libs()
    if service is None:
        service = build("drive", "v3", credentials=sign_in(cfg), cache_discovery=False)
    existing = find_existing(service, deck_path.stem, cfg.folder_id)
    if existing:
        if on_existing is None or not on_existing(existing):
            raise UploadDeclined(existing)
        target = existing[0]
        updated = (
            service.files()
            .update(
                fileId=target.file_id,
                media_body=MediaFileUpload(str(deck_path), mimetype=PPTX_MIME),
                fields="id,name,webViewLink",
                supportsAllDrives=True,
            )
            .execute()
        )
        return UploadResult(
            file_id=updated.get("id", target.file_id),
            name=updated.get("name", target.name),
            link=updated.get("webViewLink", target.link),
            action="overwritten",
        )
    created = (
        service.files()
        .create(
            body=upload_metadata(deck_path, cfg.folder_id),
            media_body=MediaFileUpload(str(deck_path), mimetype=PPTX_MIME),
            fields="id,name,webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )
    return UploadResult(
        file_id=created["id"], name=created.get("name", deck_path.stem), link=created.get("webViewLink", "")
    )
