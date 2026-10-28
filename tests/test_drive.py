"""Tests for shoutout_gen/drive.py and its CLI hook-up.

No network anywhere: the Drive service is a stub. The contract under test is
"invisible unless .env opts in" -- the generator never touches Google unless a
folder is configured, and when it is, exactly one convert-to-Slides upload goes
into that folder.
"""
from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import generate_shoutouts as cli  # noqa: E402
import upload_to_slides  # noqa: E402
from shoutout_gen import drive  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
EXAMPLE_FORM = ROOT / "samples" / "example_form.csv"
FOLDER_ID = "1AbCdEfGhIjKlMnOpQrStUvWxYz_0123456"
FOLDER_LINK = f"https://drive.google.com/drive/folders/{FOLDER_ID}?usp=sharing"
DRIVE_KEYS = (
    drive.KEY_FOLDER, drive.KEY_SWITCH, drive.KEY_CLIENT_ID, drive.KEY_CLIENT_SECRET,
    drive.KEY_CLIENT_FILE, drive.KEY_TOKEN_FILE,
)


@pytest.fixture(autouse=True)
def _clean_environment(monkeypatch):
    """The developer's own shell must not leak DRIVE_* settings into the CLI tests."""
    for key in DRIVE_KEYS:
        monkeypatch.delenv(key, raising=False)


def _env(tmp_path: Path, text: str) -> Path:
    env = tmp_path / ".env"
    env.write_text(text, encoding="utf-8")
    return env


# --------------------------------------------------------------------------- #
# .env parsing
# --------------------------------------------------------------------------- #
def test_missing_env_file_is_empty(tmp_path):
    assert drive.read_env_file(tmp_path / "nope.env") == {}


def test_env_file_syntax(tmp_path):
    env = _env(tmp_path, "\n".join([
        "# comment",
        "",
        "DRIVE_FOLDER=" + FOLDER_LINK,
        "export DRIVE_UPLOAD = on ",
        'DRIVE_CLIENT_ID="quoted # not a comment"',
        "DRIVE_CLIENT_SECRET='single'",
        "DRIVE_TOKEN_FILE=tok.json # trailing comment",
    ]))
    assert drive.read_env_file(env) == {
        "DRIVE_FOLDER": FOLDER_LINK,
        "DRIVE_UPLOAD": "on",
        "DRIVE_CLIENT_ID": "quoted # not a comment",
        "DRIVE_CLIENT_SECRET": "single",
        "DRIVE_TOKEN_FILE": "tok.json",
    }


def test_quoted_value_with_a_trailing_comment(tmp_path):
    env = _env(tmp_path, "\n".join([
        f"DRIVE_FOLDER={FOLDER_LINK}",
        'DRIVE_UPLOAD="off" # paused for break',
        "DRIVE_CLIENT_SECRET='s3cret' #x",
    ]))
    assert drive.read_env_file(env)["DRIVE_CLIENT_SECRET"] == "s3cret"
    assert drive.load_config(env, environ={}) is None  # the quoted kill switch still switches it off


def test_text_after_a_quoted_value_is_an_error(tmp_path):
    env = _env(tmp_path, 'DRIVE_FOLDER="abc" def\n')
    with pytest.raises(drive.DriveConfigError, match=r"\.env:1"):
        drive.read_env_file(env)


def test_non_utf8_env_file_is_a_config_error(tmp_path):
    env = tmp_path / ".env"
    env.write_text(f"DRIVE_FOLDER={FOLDER_ID}\n", encoding="utf-16")  # Windows PowerShell's Out-File default
    with pytest.raises(drive.DriveConfigError, match="UTF-8"):
        drive.read_env_file(env)


def test_env_file_with_bom_and_crlf(tmp_path):
    env = tmp_path / ".env"
    env.write_bytes(b"\xef\xbb\xbfDRIVE_FOLDER=" + FOLDER_ID.encode() + b"\r\n")
    assert drive.read_env_file(env) == {"DRIVE_FOLDER": FOLDER_ID}


def test_malformed_line_is_an_error_not_silently_skipped(tmp_path):
    env = _env(tmp_path, "DRIVE_FOLDER=x\nthis is not a setting\n")
    with pytest.raises(drive.DriveConfigError, match=r"\.env:2"):
        drive.read_env_file(env)


@pytest.mark.parametrize("text", [
    FOLDER_LINK,
    f"https://drive.google.com/drive/u/0/folders/{FOLDER_ID}",
    f"https://drive.google.com/drive/folders/{FOLDER_ID}",
    FOLDER_ID,
    f"  {FOLDER_ID}\n",
])
def test_extract_folder_id(text):
    assert drive.extract_folder_id(text) == FOLDER_ID


@pytest.mark.parametrize("bad", ["", "not a link", "https://drive.google.com/file/d/abcdefghijk/view", "short"])
def test_extract_folder_id_rejects_non_folders(bad):
    with pytest.raises(drive.DriveConfigError):
        drive.extract_folder_id(bad)


# --------------------------------------------------------------------------- #
# load_config: off unless opted in
# --------------------------------------------------------------------------- #
def test_off_without_env_file(tmp_path):
    assert drive.load_config(tmp_path / ".env", environ={}) is None


def test_off_without_folder(tmp_path):
    env = _env(tmp_path, "DRIVE_CLIENT_ID=abc\nDRIVE_FOLDER=\n")
    assert drive.load_config(env, environ={}) is None


@pytest.mark.parametrize("value", ["off", "OFF", "false", "0", "no"])
def test_kill_switch(tmp_path, value):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\nDRIVE_UPLOAD={value}\n")
    assert drive.load_config(env, environ={}) is None


def test_enabled_with_defaults(tmp_path):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\n")
    assert drive.load_config(env, environ={}) == drive.DriveConfig(
        folder_id=FOLDER_ID,
        token_file=tmp_path.resolve() / ".google" / "token.json",
        client_file=tmp_path.resolve() / ".google" / "credentials.json",
    )


def test_paths_resolve_against_env_file_unless_absolute(tmp_path):
    absolute = (tmp_path / "elsewhere" / "creds.json").resolve()
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_ID}\nDRIVE_TOKEN_FILE=cache/tok.json\nDRIVE_CLIENT_FILE={absolute}\n")
    cfg = drive.load_config(env, environ={})
    assert cfg.token_file == tmp_path.resolve() / "cache" / "tok.json"
    assert cfg.client_file == absolute


def test_inline_client_is_read(tmp_path):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_ID}\nDRIVE_CLIENT_ID=id.apps\nDRIVE_CLIENT_SECRET=s3cret\n")
    cfg = drive.load_config(env, environ={})
    assert (cfg.client_id, cfg.client_secret) == ("id.apps", "s3cret")


def test_environment_overrides_file(tmp_path):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_ID}\n")
    assert drive.load_config(env, environ={"DRIVE_UPLOAD": "off"}) is None
    other = "0123456789abcdef_other"
    assert drive.load_config(env, environ={"DRIVE_FOLDER": other}).folder_id == other
    assert drive.load_config(tmp_path / "absent.env", environ={"DRIVE_FOLDER": other}).folder_id == other


def test_bad_folder_link_is_a_config_error(tmp_path):
    env = _env(tmp_path, "DRIVE_FOLDER=https://example.com/not-drive\n")
    with pytest.raises(drive.DriveConfigError, match="DRIVE_FOLDER"):
        drive.load_config(env, environ={})


# --------------------------------------------------------------------------- #
# The upload request (stub service, no network)
# --------------------------------------------------------------------------- #
class _StubDrive:
    """Mimics ``service.files().list/create/update(...).execute()`` and records the requests.

    ``existing`` is what the folder already holds, as Drive would list it.
    """

    def __init__(self, existing=()):
        self.existing = list(existing)
        self.queries = []  # list() kwargs
        self.requests = []  # ("create" | "update", kwargs)
        self._pending = None

    def files(self):
        return self

    def list(self, **kwargs):
        self.queries.append(kwargs)
        self._pending = ("list", kwargs)
        return self

    def create(self, **kwargs):
        self.requests.append(("create", kwargs))
        self._pending = ("create", kwargs)
        return self

    def update(self, **kwargs):
        # As Drive v3 does: an update cannot set parents (that takes addParents/removeParents).
        assert "parents" not in kwargs.get("body", {}), "Drive rejects parents in an update body"
        self.requests.append(("update", kwargs))
        self._pending = ("update", kwargs)
        return self

    def execute(self):
        kind, kwargs = self._pending
        if kind == "list":
            return {"files": self.existing}
        if kind == "create":
            name = kwargs["body"]["name"]
            return {"id": "file123", "name": name, "webViewLink": "https://docs.google.com/presentation/d/file123/edit"}
        file_id = kwargs["fileId"]
        return {"id": file_id, "name": "GM #25 9.10---shoutouts", "webViewLink": f"https://docs.google.com/presentation/d/{file_id}/edit"}


OLD_COPY = {"id": "old1", "name": "GM #25 9.10---shoutouts", "webViewLink": "https://docs.google.com/presentation/d/old1/edit", "modifiedTime": "2026-09-10T20:00:00.000Z"}
NEWER_COPY = {"id": "new2", "name": "GM #25 9.10---shoutouts", "webViewLink": "https://docs.google.com/presentation/d/new2/edit", "modifiedTime": "2026-09-11T08:30:00.000Z"}


def _cfg(tmp_path):
    return drive.DriveConfig(folder_id=FOLDER_ID, token_file=tmp_path / "t.json", client_file=tmp_path / "c.json")


def test_upload_converts_into_the_folder(tmp_path):
    pytest.importorskip("googleapiclient")
    deck = tmp_path / "GM #25 9.10---shoutouts.pptx"
    deck.write_bytes(b"not really a deck")
    stub = _StubDrive()

    result = drive.upload_deck(deck, _cfg(tmp_path), service=stub)

    (query,) = stub.queries
    assert query["q"] == drive.existing_query("GM #25 9.10---shoutouts", FOLDER_ID)
    assert query["supportsAllDrives"] is True and query["includeItemsFromAllDrives"] is True
    ((kind, request),) = stub.requests
    assert kind == "create"
    assert request["body"] == {
        "name": "GM #25 9.10---shoutouts", "parents": [FOLDER_ID], "mimeType": drive.GOOGLE_SLIDES_MIME,
    }
    assert request["media_body"].mimetype() == drive.PPTX_MIME
    assert request["supportsAllDrives"] is True
    assert "webViewLink" in request["fields"]
    assert result == drive.UploadResult(
        "file123", "GM #25 9.10---shoutouts", "https://docs.google.com/presentation/d/file123/edit"
    )
    assert result.action == "created"


def test_upload_missing_deck_fails_before_touching_drive(tmp_path):
    stub = _StubDrive()
    with pytest.raises(FileNotFoundError):
        drive.upload_deck(tmp_path / "missing.pptx", _cfg(tmp_path), service=stub)
    assert stub.requests == [] and stub.queries == []


# --------------------------------------------------------------------------- #
# A same-name file is already in the folder: overwrite in place, or leave it
# --------------------------------------------------------------------------- #
def _deck(tmp_path):
    deck = tmp_path / "GM #25 9.10---shoutouts.pptx"
    deck.write_bytes(b"not really a deck")
    return deck


def test_existing_query_escapes_the_name():
    q = drive.existing_query("it's a \\ name", FOLDER_ID)
    assert "name = 'it\\'s a \\\\ name'" in q
    assert f"'{FOLDER_ID}' in parents" in q and "trashed = false" in q and drive.GOOGLE_SLIDES_MIME in q


def test_find_existing_lists_newest_first(tmp_path):
    stub = _StubDrive(existing=[OLD_COPY, NEWER_COPY])
    found = drive.find_existing(stub, "GM #25 9.10---shoutouts", FOLDER_ID)
    assert stub.queries[0]["orderBy"] == "modifiedTime desc"  # the newest is in the page however many copies
    assert [f.file_id for f in found] == ["new2", "old1"]
    assert found[0] == drive.ExistingFile("new2", "GM #25 9.10---shoutouts", NEWER_COPY["webViewLink"], NEWER_COPY["modifiedTime"])


def test_overwrite_updates_the_newest_copy_in_place(tmp_path):
    pytest.importorskip("googleapiclient")
    stub = _StubDrive(existing=[OLD_COPY, NEWER_COPY])
    seen = []

    def yes(existing):
        seen.append([f.file_id for f in existing])
        return True

    result = drive.upload_deck(_deck(tmp_path), _cfg(tmp_path), service=stub, on_existing=yes)

    assert seen == [["new2", "old1"]]
    ((kind, request),) = stub.requests
    assert kind == "update" and request["fileId"] == "new2"
    assert "body" not in request  # no metadata: the file keeps its name, folder and Slides type
    assert request["media_body"].mimetype() == drive.PPTX_MIME
    assert request["supportsAllDrives"] is True and "webViewLink" in request["fields"]
    assert result == drive.UploadResult("new2", "GM #25 9.10---shoutouts", NEWER_COPY["webViewLink"], action="overwritten")


def test_declined_overwrite_touches_nothing(tmp_path):
    pytest.importorskip("googleapiclient")
    stub = _StubDrive(existing=[OLD_COPY])
    with pytest.raises(drive.UploadDeclined) as exc:
        drive.upload_deck(_deck(tmp_path), _cfg(tmp_path), service=stub, on_existing=lambda existing: False)
    assert exc.value.existing[0].file_id == "old1"
    assert "GM #25 9.10---shoutouts" in str(exc.value)
    assert stub.requests == []


def test_no_callback_never_overwrites(tmp_path):
    pytest.importorskip("googleapiclient")
    stub = _StubDrive(existing=[OLD_COPY])
    with pytest.raises(drive.UploadDeclined):
        drive.upload_deck(_deck(tmp_path), _cfg(tmp_path), service=stub)
    assert stub.requests == []


def test_callback_not_consulted_when_nothing_exists(tmp_path):
    pytest.importorskip("googleapiclient")
    stub = _StubDrive()
    result = drive.upload_deck(_deck(tmp_path), _cfg(tmp_path), service=stub, on_existing=lambda existing: pytest.fail("asked"))
    assert result.action == "created"


@pytest.mark.parametrize("answer, expected", [("y", True), ("Y", True), ("yes", True), (" yes ", True), ("n", False), ("", False), ("maybe", False)])
def test_terminal_prompt_only_yes_overwrites(monkeypatch, capsys, answer, expected):
    prompts = []
    monkeypatch.setattr("builtins.input", lambda text: prompts.append(text) or answer)
    existing = [drive.ExistingFile("new2", "GM #25 9.10---shoutouts", NEWER_COPY["webViewLink"], NEWER_COPY["modifiedTime"]),
                drive.ExistingFile("old1", "GM #25 9.10---shoutouts", OLD_COPY["webViewLink"], OLD_COPY["modifiedTime"])]
    assert drive.confirm_on_terminal(existing) is expected
    (prompt,) = prompts
    assert "GM #25 9.10---shoutouts" in prompt and "2 copies" in prompt and "2026-09-11 08:30 UTC" in prompt and "[y/N]" in prompt


def test_terminal_prompt_without_a_terminal_keeps_the_file(monkeypatch, capsys):
    def no_stdin(text):
        raise EOFError

    monkeypatch.setattr("builtins.input", no_stdin)
    existing = [drive.ExistingFile("old1", "GM #25 9.10---shoutouts", OLD_COPY["webViewLink"], OLD_COPY["modifiedTime"])]
    assert drive.confirm_on_terminal(existing) is False
    assert "--overwrite" in capsys.readouterr().err


def test_always_overwrite():
    assert drive.always_overwrite([]) is True


def test_client_config_shape():
    conf = drive.client_config("id", "secret")["installed"]
    assert conf["client_id"] == "id" and conf["client_secret"] == "secret"
    assert conf["redirect_uris"] == ["http://localhost"]


# --------------------------------------------------------------------------- #
# CLI gating
# --------------------------------------------------------------------------- #
class _Calls(list):
    """The recorded upload calls, plus the on_existing callback the CLI handed over."""

    on_existing = None


@pytest.fixture
def upload_spy(monkeypatch):
    """Replace the real upload with a recorder; the tests assert on what reached it."""
    calls = _Calls()

    def fake_upload(deck_path, cfg, service=None, on_existing=None):
        calls.append((Path(deck_path), cfg))
        calls.on_existing = on_existing
        return drive.UploadResult("f1", Path(deck_path).stem, "https://docs.google.com/presentation/d/f1/edit")

    monkeypatch.setattr(drive, "upload_deck", fake_upload)
    return calls


def _generate(tmp_path, env_file, *extra):
    return cli._main([
        str(EXAMPLE_FORM), "--gm", "1", "--date", "9.17", "--out", str(tmp_path / "out"),
        "--env-file", str(env_file), *extra,
    ])


def test_cli_does_not_upload_without_env(tmp_path, upload_spy, capsys):
    assert _generate(tmp_path, tmp_path / "absent.env") == 0
    assert upload_spy == []
    assert "Drive" not in capsys.readouterr().out


def test_cli_uploads_when_env_enables_it(tmp_path, upload_spy, capsys):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\n")
    assert _generate(tmp_path, env) == 0
    ((deck, cfg),) = upload_spy
    assert deck == tmp_path / "out" / "GM #1 9.17---shoutouts.pptx" and deck.exists()
    assert cfg.folder_id == FOLDER_ID
    assert "https://docs.google.com/presentation/d/f1/edit" in capsys.readouterr().out


def test_cli_no_upload_flag_wins_over_env(tmp_path, upload_spy):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\n")
    assert _generate(tmp_path, env, "--no-upload") == 0
    assert upload_spy == []


def test_cli_kill_switch_in_env(tmp_path, upload_spy):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\nDRIVE_UPLOAD=off\n")
    assert _generate(tmp_path, env) == 0
    assert upload_spy == []


def test_cli_bad_env_fails_before_generating(tmp_path, upload_spy, capsys):
    env = _env(tmp_path, "DRIVE_FOLDER=https://example.com/nope\n")
    assert _generate(tmp_path, env) == 2
    assert not (tmp_path / "out").exists()
    assert upload_spy == []
    assert "DRIVE_FOLDER" in capsys.readouterr().err


def test_cli_upload_failure_keeps_deck_and_reports(tmp_path, monkeypatch, capsys):
    def boom(deck_path, cfg, service=None, on_existing=None):
        raise RuntimeError("quota exceeded")

    monkeypatch.setattr(drive, "upload_deck", boom)
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\n")
    assert _generate(tmp_path, env) == 3
    assert (tmp_path / "out" / "GM #1 9.17---shoutouts.pptx").exists()
    err = capsys.readouterr().err
    assert "quota exceeded" in err and "upload_to_slides.py" in err


def test_cli_asks_on_the_terminal_by_default(tmp_path, upload_spy):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\n")
    assert _generate(tmp_path, env) == 0
    assert upload_spy.on_existing is drive.confirm_on_terminal


def test_cli_overwrite_flag_answers_yes(tmp_path, upload_spy):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\n")
    assert _generate(tmp_path, env, "--overwrite") == 0
    assert upload_spy.on_existing is drive.always_overwrite


def test_cli_reports_an_overwrite(tmp_path, monkeypatch, capsys):
    def overwrote(deck_path, cfg, service=None, on_existing=None):
        return drive.UploadResult("old1", Path(deck_path).stem, "https://docs.google.com/presentation/d/old1/edit", action="overwritten")

    monkeypatch.setattr(drive, "upload_deck", overwrote)
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\n")
    assert _generate(tmp_path, env) == 0
    out = capsys.readouterr().out
    assert "overwrote" in out and "https://docs.google.com/presentation/d/old1/edit" in out


def test_cli_declined_overwrite_keeps_deck_and_says_how_to_replace(tmp_path, monkeypatch, capsys):
    def declined(deck_path, cfg, service=None, on_existing=None):
        raise drive.UploadDeclined([drive.ExistingFile("old1", Path(deck_path).stem, "https://docs.google.com/presentation/d/old1/edit", "")])

    monkeypatch.setattr(drive, "upload_deck", declined)
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_LINK}\n")
    assert _generate(tmp_path, env) == 0
    assert (tmp_path / "out" / "GM #1 9.17---shoutouts.pptx").exists()
    out = capsys.readouterr().out
    assert "not uploaded" in out and "old1" in out and "--overwrite" in out


# --------------------------------------------------------------------------- #
# upload_to_slides.py (retry path)
# --------------------------------------------------------------------------- #
def test_retry_script_requires_enabled_config(tmp_path, upload_spy, capsys):
    deck = tmp_path / "d.pptx"
    deck.write_bytes(b"x")
    assert upload_to_slides._main([str(deck), "--env-file", str(tmp_path / "absent.env")]) == 2
    assert upload_spy == []
    assert "DRIVE_FOLDER" in capsys.readouterr().err


def test_retry_script_uploads_existing_deck(tmp_path, upload_spy, capsys):
    deck = tmp_path / "GM #3 10.1---shoutouts.pptx"
    deck.write_bytes(b"x")
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_ID}\n")
    assert upload_to_slides._main([str(deck), "--env-file", str(env)]) == 0
    assert [d for d, _ in upload_spy] == [deck]
    assert upload_spy.on_existing is drive.confirm_on_terminal
    assert "docs.google.com" in capsys.readouterr().out


def test_retry_script_overwrite_flag_and_decline(tmp_path, upload_spy, monkeypatch, capsys):
    deck = tmp_path / "GM #3 10.1---shoutouts.pptx"
    deck.write_bytes(b"x")
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_ID}\n")
    assert upload_to_slides._main([str(deck), "--env-file", str(env), "--overwrite"]) == 0
    assert upload_spy.on_existing is drive.always_overwrite

    def declined(deck_path, cfg, service=None, on_existing=None):
        raise drive.UploadDeclined([drive.ExistingFile("old1", deck.stem, "https://docs.google.com/presentation/d/old1/edit", "")])

    monkeypatch.setattr(drive, "upload_deck", declined)
    assert upload_to_slides._main([str(deck), "--env-file", str(env)]) == 0
    assert "not uploaded" in capsys.readouterr().out


# --------------------------------------------------------------------------- #
# load_signin: the sign-in without the upload (shared with the Google Sheet source)
# --------------------------------------------------------------------------- #
def test_load_signin_defaults_without_env(tmp_path):
    assert drive.load_signin(tmp_path / "absent.env", environ={}) == drive.SignInConfig(
        token_file=tmp_path.resolve() / ".google" / "token.json",
        client_file=tmp_path.resolve() / ".google" / "credentials.json",
    )


def test_load_signin_does_not_need_a_folder(tmp_path):
    env = _env(tmp_path, "DRIVE_CLIENT_ID=id.apps\nDRIVE_CLIENT_SECRET=s3cret\nDRIVE_TOKEN_FILE=tok.json\n")
    cfg = drive.load_signin(env, environ={})
    assert (cfg.client_id, cfg.client_secret) == ("id.apps", "s3cret")
    assert cfg.token_file == tmp_path.resolve() / "tok.json"
    assert drive.load_config(env, environ={}) is None  # no folder: upload still off


def test_load_config_carries_the_same_signin(tmp_path):
    env = _env(tmp_path, f"DRIVE_FOLDER={FOLDER_ID}\nDRIVE_CLIENT_ID=id.apps\nDRIVE_TOKEN_FILE=tok.json\n")
    cfg, signin = drive.load_config(env, environ={}), drive.load_signin(env, environ={})
    assert (cfg.token_file, cfg.client_file, cfg.client_id) == (signin.token_file, signin.client_file, signin.client_id)


# --------------------------------------------------------------------------- #
# sign_in: cached token -> refresh -> consent (Google's classes faked, no browser)
# --------------------------------------------------------------------------- #
class _FakeCreds:
    def __init__(self, valid, expired=False, refresh_token=None, refresh_fails=False):
        self.valid, self.expired, self.refresh_token, self._refresh_fails = valid, expired, refresh_token, refresh_fails

    def refresh(self, request):
        from google.auth.exceptions import RefreshError

        if self._refresh_fails:
            raise RefreshError("token has been revoked")
        self.valid = True

    def to_json(self):
        return '{"token": "fresh"}'


@pytest.fixture
def fake_google(monkeypatch):
    """The cached token sign_in finds (``.cached``) and the consents it asks for (``.consents``)."""
    pytest.importorskip("google_auth_oauthlib")
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow

    state = SimpleNamespace(cached=None, consents=[])

    class _Flow:
        def __init__(self, source):
            self.source = source

        def run_local_server(self, port):
            state.consents.append(self.source)
            return _FakeCreds(valid=True)

    monkeypatch.setattr(Credentials, "from_authorized_user_file", lambda path, scopes: state.cached)
    monkeypatch.setattr(InstalledAppFlow, "from_client_config", lambda config, scopes: _Flow(("inline", config)))
    monkeypatch.setattr(InstalledAppFlow, "from_client_secrets_file", lambda path, scopes: _Flow(("file", path)))
    return state


def test_sign_in_reuses_a_valid_cached_token(tmp_path, fake_google):
    token = tmp_path / "t.json"
    token.write_text("{}", encoding="utf-8")
    fake_google.cached = _FakeCreds(valid=True)
    assert drive.sign_in(drive.SignInConfig(token_file=token, client_file=tmp_path / "c.json")) is fake_google.cached
    assert fake_google.consents == []


def test_sign_in_consents_with_the_inline_client_and_caches_the_token(tmp_path, fake_google):
    client_file = tmp_path / "c.json"
    client_file.write_text("{}", encoding="utf-8")  # present, but the inline client wins
    cfg = drive.SignInConfig(
        token_file=tmp_path / "cache" / "t.json", client_file=client_file, client_id="id.apps", client_secret="s3cret",
    )
    drive.sign_in(cfg)
    assert fake_google.consents == [("inline", drive.client_config("id.apps", "s3cret"))]
    assert cfg.token_file.read_text(encoding="utf-8") == '{"token": "fresh"}'


def test_sign_in_asks_again_when_the_refresh_is_refused(tmp_path, fake_google):
    token, client_file = tmp_path / "t.json", tmp_path / "c.json"
    token.write_text("{}", encoding="utf-8")
    client_file.write_text("{}", encoding="utf-8")
    fake_google.cached = _FakeCreds(valid=False, expired=True, refresh_token="r", refresh_fails=True)
    drive.sign_in(drive.SignInConfig(token_file=token, client_file=client_file))
    assert fake_google.consents == [("file", str(client_file))]
    assert token.read_text(encoding="utf-8") == '{"token": "fresh"}'
