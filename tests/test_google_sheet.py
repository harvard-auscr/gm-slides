"""Google Sheet source: link parsing, the CSV download (stub session, no network), and the CLI hook-up.

The contract under test: a sheet link -- on the command line or as SHOUTOUT_SHEET
in .env -- is pulled into input/ as the CSV the reader already understands, with
the upload's sign-in, and nothing is generated if the pull fails. --since is
tested here too, because the year-long form is what makes it necessary.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import generate_shoutouts as cli  # noqa: E402
from shoutout_gen import drive, google_sheet  # noqa: E402
from shoutout_gen.google_sheet import SheetDownloadError, SheetLinkError, SheetRef  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
EXAMPLE_FORM = ROOT / "samples" / "example_form.csv"
SHEET_ID = "1Zx9SyntheticSheetId_ForTests-0123456789abc"
GID = "987654321"
# Exactly what the address bar gives when a tab is open.
SHEET_LINK = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?resourcekey=&gid={GID}#gid={GID}"
CSV_HEADER = b"Timestamp,Shout-out\r\n"


@pytest.fixture(autouse=True)
def _clean_environment(monkeypatch):
    """The developer's own shell must not leak Google settings into the CLI tests."""
    for key in (google_sheet.KEY_SHEET, drive.KEY_FOLDER, drive.KEY_SWITCH, drive.KEY_CLIENT_ID,
                drive.KEY_CLIENT_SECRET, drive.KEY_CLIENT_FILE, drive.KEY_TOKEN_FILE):
        monkeypatch.delenv(key, raising=False)


def _env(tmp_path: Path, text: str) -> Path:
    env = tmp_path / ".env"
    env.write_text(text, encoding="utf-8")
    return env


# --------------------------------------------------------------------------- #
# Links
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize("text, expected", [
    (SHEET_LINK, SheetRef(SHEET_ID, GID)),
    (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit#gid=42", SheetRef(SHEET_ID, "42")),
    (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?gid=7", SheetRef(SHEET_ID, "7")),
    (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?usp=sharing", SheetRef(SHEET_ID, None)),
    (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}", SheetRef(SHEET_ID, None)),
    (f"https://docs.google.com/spreadsheets/u/1/d/{SHEET_ID}/edit#gid=0", SheetRef(SHEET_ID, "0")),
    (f"  https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit#gid={GID}\n", SheetRef(SHEET_ID, GID)),
])
def test_parse_sheet_link(text, expected):
    assert google_sheet.is_sheet_link(text)
    assert google_sheet.parse_sheet_link(text) == expected


@pytest.mark.parametrize("bad", [
    "",
    "input/form.csv",
    r"C:\club\input\GM #25 9.10---form.csv",
    "https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrStUvWxYz_0123456",
    "https://docs.google.com/document/d/1AbCdEfGhIjKlMnOpQrStUvWxYz_0123456/edit",
    SHEET_ID,  # a bare ID is not accepted: it would be indistinguishable from a file name
])
def test_not_a_sheet_link(bad):
    assert not google_sheet.is_sheet_link(bad)
    with pytest.raises(SheetLinkError):
        google_sheet.parse_sheet_link(bad)


@pytest.mark.parametrize("ref", [SheetRef(SHEET_ID, GID), SheetRef(SHEET_ID, None)])
def test_ref_link_round_trips(ref):
    assert google_sheet.parse_sheet_link(ref.link) == ref


def test_form_name_matches_the_deck_naming():
    assert google_sheet.form_name_for(26, "9.24") == "GM #26 9.24---form.csv"


# --------------------------------------------------------------------------- #
# .env
# --------------------------------------------------------------------------- #
def test_sheet_link_from_env_file(tmp_path):
    env = _env(tmp_path, f"DRIVE_FOLDER=\nSHOUTOUT_SHEET={SHEET_LINK}\n")
    assert google_sheet.load_sheet_link(env, environ={}) == SHEET_LINK


def test_sheet_link_unset_or_absent(tmp_path):
    assert google_sheet.load_sheet_link(tmp_path / "absent.env", environ={}) == ""
    assert google_sheet.load_sheet_link(_env(tmp_path, "SHOUTOUT_SHEET=\n"), environ={}) == ""


def test_environment_overrides_env_file_link(tmp_path):
    env = _env(tmp_path, f"SHOUTOUT_SHEET={SHEET_LINK}\n")
    other = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit#gid=1"
    assert google_sheet.load_sheet_link(env, environ={"SHOUTOUT_SHEET": other}) == other


# --------------------------------------------------------------------------- #
# The download (stub session, no network)
# --------------------------------------------------------------------------- #
class _Response:
    def __init__(self, status=200, content_type="text/csv; charset=utf-8", content=CSV_HEADER):
        self.status_code = status
        self.headers = {"content-type": content_type}
        self.content = content


class _StubSession:
    """Mimics ``requests.Session.get`` and records the request.

    ``responses`` are handed out in order; the last one repeats.
    """

    def __init__(self, *responses):
        self.responses = list(responses) or [_Response()]
        self.calls = []

    def get(self, url, params=None):
        self.calls.append((url, dict(params or {})))
        return self.responses.pop(0) if len(self.responses) > 1 else self.responses[0]


@pytest.fixture(autouse=True)
def _no_retry_delay(monkeypatch):
    monkeypatch.setattr(google_sheet, "_RETRY_DELAY_S", 0)


def _signin(tmp_path):
    return drive.SignInConfig(token_file=tmp_path / "t.json", client_file=tmp_path / "c.json")


def test_fetch_requests_the_tab_as_csv(tmp_path):
    session = _StubSession()
    data = google_sheet.fetch_csv(SheetRef(SHEET_ID, GID), _signin(tmp_path), session=session)
    assert data == CSV_HEADER
    assert session.calls == [(f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export", {"format": "csv", "gid": GID})]


def test_fetch_without_a_tab_leaves_gid_out(tmp_path):
    session = _StubSession()
    google_sheet.fetch_csv(SheetRef(SHEET_ID, None), _signin(tmp_path), session=session)
    assert session.calls[0][1] == {"format": "csv"}


@pytest.mark.parametrize("status, content_type, hint", [
    (400, "text/html; charset=utf-8", "tab (gid)"),
    (401, "text/html; charset=utf-8", "sign-in"),
    (403, "text/html; charset=utf-8", "share it"),
    (404, "text/html; charset=utf-8", "no spreadsheet"),
    (200, "text/html; charset=utf-8", "other than the tab's CSV"),  # a login page with a 200
    (500, "", "other than the tab's CSV"),
])
def test_fetch_rejects_anything_but_csv(tmp_path, status, content_type, hint):
    session = _StubSession(_Response(status, content_type, b"<!DOCTYPE html>"))
    with pytest.raises(SheetDownloadError) as exc:
        google_sheet.fetch_csv(SheetRef(SHEET_ID, GID), _signin(tmp_path), session=session)
    message = str(exc.value)
    assert f"HTTP {status}" in message and hint in message and SheetRef(SHEET_ID, GID).link in message
    assert len(session.calls) == 2  # one retry, then give up


def test_fetch_retries_a_passing_failure_once(tmp_path):
    session = _StubSession(_Response(400, "text/html", b"<!DOCTYPE html>"), _Response())
    assert google_sheet.fetch_csv(SheetRef(SHEET_ID, GID), _signin(tmp_path), session=session) == CSV_HEADER
    assert len(session.calls) == 2 and session.calls[0] == session.calls[1]


def test_download_writes_the_bytes_and_creates_the_folder(tmp_path):
    dest = tmp_path / "in" / "GM #26 9.24---form.csv"
    body = b"\xef\xbb\xbfTimestamp,Shout-out\r\n9/23/2026 16:35:52,thanks\r\n"
    out = google_sheet.download(SheetRef(SHEET_ID, GID), _signin(tmp_path), dest, session=_StubSession(_Response(content=body)))
    assert out == dest and dest.read_bytes() == body  # verbatim: the BOM and CRLF are the reader's business


def test_download_failure_leaves_no_file(tmp_path):
    dest = tmp_path / "in" / "form.csv"
    with pytest.raises(SheetDownloadError):
        google_sheet.download(SheetRef(SHEET_ID, GID), _signin(tmp_path), dest, session=_StubSession(_Response(403, "text/html")))
    assert not dest.exists()


# --------------------------------------------------------------------------- #
# CLI: link -> input/ CSV -> deck
# --------------------------------------------------------------------------- #
@pytest.fixture
def download_spy(monkeypatch):
    """Replace the real download with one that drops the sample form where the CLI asked."""
    calls = []

    def fake_download(ref, signin, dest, session=None):
        calls.append((ref, signin, Path(dest)))
        Path(dest).parent.mkdir(parents=True, exist_ok=True)
        Path(dest).write_bytes(EXAMPLE_FORM.read_bytes())
        return Path(dest)

    monkeypatch.setattr(google_sheet, "download", fake_download)
    return calls


def _generate(tmp_path, env_file, *extra):
    return cli._main([
        "--gm", "1", "--date", "9.17", "--out", str(tmp_path / "out"), "--input-dir", str(tmp_path / "in"),
        "--env-file", str(env_file), "--no-upload", *extra,
    ])


def test_cli_link_argument_pulls_then_generates(tmp_path, download_spy, capsys):
    env = _env(tmp_path, "DRIVE_TOKEN_FILE=tok.json\n")
    assert _generate(tmp_path, env, SHEET_LINK) == 0
    ((ref, signin, dest),) = download_spy
    assert ref == SheetRef(SHEET_ID, GID)
    assert signin.token_file == tmp_path.resolve() / "tok.json"  # the upload's sign-in, no folder needed
    assert dest == tmp_path / "in" / "GM #1 9.17---form.csv" and dest.exists()
    assert (tmp_path / "out" / "GM #1 9.17---shoutouts.pptx").exists()
    out = capsys.readouterr().out
    assert "pulled" in out and str(dest) in out and "submitted" in out


def test_cli_link_from_env_when_no_argument(tmp_path, download_spy):
    env = _env(tmp_path, f"SHOUTOUT_SHEET={SHEET_LINK}\n")
    assert _generate(tmp_path, env) == 0
    assert [ref for ref, _, _ in download_spy] == [SheetRef(SHEET_ID, GID)]


def test_cli_file_argument_beats_env_link(tmp_path, download_spy):
    env = _env(tmp_path, f"SHOUTOUT_SHEET={SHEET_LINK}\n")
    assert _generate(tmp_path, env, str(EXAMPLE_FORM)) == 0
    assert download_spy == []


def test_cli_nothing_to_read(tmp_path, download_spy, capsys):
    assert _generate(tmp_path, tmp_path / "absent.env") == 2
    assert download_spy == []
    assert not (tmp_path / "out").exists()
    assert google_sheet.KEY_SHEET in capsys.readouterr().err


def test_cli_download_failure_exits_before_generating(tmp_path, monkeypatch, capsys):
    def boom(ref, signin, dest, session=None):
        raise SheetDownloadError("could not download: HTTP 403")

    monkeypatch.setattr(google_sheet, "download", boom)
    assert _generate(tmp_path, tmp_path / "absent.env", SHEET_LINK) == 2
    assert not (tmp_path / "out").exists()
    err = capsys.readouterr().err
    assert "HTTP 403" in err and "Download" in err


def test_cli_malformed_env_fails_even_with_no_upload(tmp_path, download_spy, capsys):
    env = _env(tmp_path, "this is not a setting\n")
    assert _generate(tmp_path, env, SHEET_LINK) == 2
    assert download_spy == []
    assert ".env:1" in capsys.readouterr().err


# --------------------------------------------------------------------------- #
# CLI: --since
# --------------------------------------------------------------------------- #
def _two_weeks(tmp_path):
    p = tmp_path / "two_weeks.csv"
    p.write_text(
        "Timestamp,Shout-out\n"
        "9/10/2026 19:00:00,old news\n"
        "9/17/2026 21:00:00,fresh\n"
        "9/20/2026 12:00:00,also fresh\n",
        encoding="utf-8",
    )
    return p


def test_cli_since_keeps_only_newer_rows(tmp_path, capsys):
    assert _generate(tmp_path, tmp_path / "absent.env", str(_two_weeks(tmp_path)), "--since", "2026-09-17") == 0
    out = capsys.readouterr().out
    assert "2 kept, 1 older skipped" in out and "2 shout-outs" in out
    assert "submitted 09/17/2026 21:00 -> 09/20/2026 12:00" in out


def test_cli_since_with_nothing_left(tmp_path, capsys):
    assert _generate(tmp_path, tmp_path / "absent.env", str(_two_weeks(tmp_path)), "--since", "2026-12-01") == 2
    assert not (tmp_path / "out").exists()
    assert "no shout-outs since 12/01/2026" in capsys.readouterr().err


def test_cli_rejects_a_bad_since(tmp_path):
    with pytest.raises(SystemExit) as exc:
        _generate(tmp_path, tmp_path / "absent.env", str(_two_weeks(tmp_path)), "--since", "yesterday")
    assert exc.value.code == 2


# --------------------------------------------------------------------------- #
# CLI: placeholder rows
# --------------------------------------------------------------------------- #
def test_cli_skips_placeholder_rows_and_says_so(tmp_path, capsys):
    p = tmp_path / "with_slashes.csv"
    p.write_text(
        "Timestamp,Shout-out,QOTW\n"
        "9/23/2026 20:08:35,/,Casey Moore\n"
        "9/23/2026 20:08:38,/,Casey Moore\n"
        "9/23/2026 20:09:00,thanks for the snacks,someone\n",
        encoding="utf-8",
    )
    assert _generate(tmp_path, tmp_path / "absent.env", str(p)) == 0
    out = capsys.readouterr().out
    assert "skipped 2 placeholder row(s)" in out and "'/'" in out and "1 shout-outs" in out


def test_cli_only_placeholders_is_an_empty_sheet(tmp_path, capsys):
    p = tmp_path / "only_slashes.csv"
    p.write_text("Timestamp,Shout-out\n9/23/2026 20:08:35,/\n", encoding="utf-8")
    assert _generate(tmp_path, tmp_path / "absent.env", str(p)) == 2
    assert "no non-empty shout-outs" in capsys.readouterr().err
