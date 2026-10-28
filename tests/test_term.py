"""Terminal colour: on for a console, off for anything piped or captured."""
from __future__ import annotations

import io
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import generate_shoutouts  # noqa: E402
from shoutout_gen import term  # noqa: E402

ESC = "\x1b["


@pytest.fixture(autouse=True)
def _no_colour_env(monkeypatch):
    monkeypatch.delenv("NO_COLOR", raising=False)
    monkeypatch.delenv("FORCE_COLOR", raising=False)


class _Tty(io.StringIO):
    """Claims to be a terminal; ``fileno`` is the file behind it, if any."""

    def __init__(self, fd: int | None = None):
        super().__init__()
        self._fd = fd

    def isatty(self) -> bool:
        return True

    def fileno(self) -> int:
        if self._fd is None:
            raise io.UnsupportedOperation("fileno")
        return self._fd


def test_plain_unless_the_stream_is_a_terminal():
    assert term.paint("hi", "red", stream=io.StringIO()) == "hi"
    assert term.paint("hi", "red", stream=_Tty()) == "hi"  # no file behind it to ask


@pytest.mark.skipif(sys.platform != "win32", reason="the console check is Windows-only")
def test_a_terminal_flag_over_a_plain_file_stays_plain(tmp_path):
    with (tmp_path / "f.txt").open("w") as fh:
        assert term.paint("hi", "red", stream=_Tty(fh.fileno())) == "hi"


def test_force_color_and_no_color(monkeypatch):
    monkeypatch.setenv("FORCE_COLOR", "1")
    assert term.paint("hi", "bold", "red", stream=io.StringIO()) == "\x1b[1;31mhi\x1b[0m"
    monkeypatch.setenv("NO_COLOR", "1")
    assert term.paint("hi", "red", stream=io.StringIO()) == "hi"
    monkeypatch.delenv("NO_COLOR")
    monkeypatch.setenv("FORCE_COLOR", "0")
    assert term.paint("hi", "red", stream=io.StringIO()) == "hi"


def test_forced_colour_still_switches_the_console_to_ansi(monkeypatch):
    asked = []
    monkeypatch.setattr(term, "_console_takes_ansi", lambda fd: asked.append(fd) or False)
    monkeypatch.setenv("FORCE_COLOR", "1")
    assert term.enabled(_Tty(7))
    assert asked == [7]  # without this a plain Windows console prints the raw escape codes


def test_overwrite_prompt_is_plain_when_stderr_is_not_a_terminal(monkeypatch):
    from shoutout_gen import drive

    prompts = []
    monkeypatch.setattr(term, "enabled", lambda stream: stream is sys.stdout)
    monkeypatch.setattr("builtins.input", lambda prompt: prompts.append(prompt) or "n")
    assert not drive.confirm_on_terminal([drive.ExistingFile("f1", "GM #1", "https://x", "")])
    assert prompts and ESC not in prompts[0]


def _run(tmp_path, *rows: str) -> int:
    sheet = tmp_path / "form.csv"
    sheet.write_text("Timestamp,Shout-out\n" + "\n".join(rows) + "\n", encoding="utf-8")
    # --no-upload: DRIVE_* variables in the developer's shell would otherwise enable a real upload.
    return generate_shoutouts._main([
        str(sheet), "--gm", "1", "--date", "9.1", "--out", str(tmp_path / "out"),
        "--env-file", str(tmp_path / "no.env"), "--since", "9.2", "--no-upload",
    ])


def test_captured_output_has_no_escape_codes(tmp_path, capsys):
    assert _run(tmp_path, "9/1/2026 10:00:00,old", "9/3/2026 10:00:00,thanks") == 0
    assert _run(tmp_path, "9/1/2026 10:00:00,old") == 2
    out, err = capsys.readouterr()
    assert "1 shout-outs ->" in out and "error: no shout-outs since" in err
    assert ESC not in out + err


def test_forced_colour_marks_results_details_and_errors(tmp_path, capsys, monkeypatch):
    monkeypatch.setenv("FORCE_COLOR", "1")
    assert _run(tmp_path, "9/1/2026 10:00:00,old", "9/3/2026 10:00:00,thanks") == 0
    out = capsys.readouterr().out
    assert "\x1b[1;32m1 shout-outs ->\x1b[0m \x1b[36m" in out  # the deck line in green, its path cyan
    assert "\x1b[90msince 09/02/2026 00:00: 1 kept, 1 older skipped\x1b[0m" in out
    assert _run(tmp_path, "9/1/2026 10:00:00,old") == 2
    assert capsys.readouterr().err.startswith("\x1b[1;31merror:\x1b[0m no shout-outs since")
