"""Colour for the scripts' terminal output.

The scheme: errors red, notes yellow, the files a run made green with their
paths and links cyan, the QOTW tally bold, and the bookkeeping lines (rows
skipped, the --since window, submission times) grey so the results stand out.

Colour goes on only when the stream is a terminal, so piped, redirected and
captured output (logs, tests) stays plain text. ``NO_COLOR`` turns it off and
``FORCE_COLOR`` turns it on regardless (https://no-color.org).
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from typing import TextIO

_CODES = {"bold": 1, "red": 31, "green": 32, "yellow": 33, "magenta": 35, "cyan": 36, "grey": 90}


def _console_takes_ansi(fd: int) -> bool:
    """True if the console behind ``fd`` interprets ANSI escapes, switching that on where Windows can."""
    if sys.platform != "win32":
        return True
    import ctypes
    import msvcrt

    enable_vt = 0x0004  # ENABLE_VIRTUAL_TERMINAL_PROCESSING; off by default in a plain console window
    kernel32 = ctypes.windll.kernel32
    handle = msvcrt.get_osfhandle(fd)
    mode = ctypes.c_ulong()
    if not kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
        return False
    return bool(mode.value & enable_vt or kernel32.SetConsoleMode(handle, mode.value | enable_vt))


def enabled(stream: TextIO) -> bool:
    """Whether text written to ``stream`` gets colour."""
    if os.environ.get("NO_COLOR"):
        return False
    try:
        # Asked even when colour is forced: on Windows this is what switches the console to ANSI.
        console = stream.isatty() and _console_takes_ansi(stream.fileno())
    except (AttributeError, OSError, ValueError):  # a stand-in stream with no file behind it
        console = False
    return console or os.environ.get("FORCE_COLOR", "") not in ("", "0")


def paint(text: str, *styles: str, stream: TextIO | None = None) -> str:
    """``text`` in the named ``styles`` if ``stream`` (default stdout) takes colour, else as is.

    Each call resets at its end, so build a line from painted pieces side by
    side rather than nesting one inside another.
    """
    if not styles or not enabled(stream or sys.stdout):
        return text
    return f"\x1b[{';'.join(str(_CODES[s]) for s in styles)}m{text}\x1b[0m"


def path(value: Path | str) -> str:
    """A file path or link, in cyan."""
    return paint(str(value), "cyan")


def error(message: str) -> None:
    """``error: <message>`` on stderr."""
    print(f"{paint('error:', 'bold', 'red', stream=sys.stderr)} {message}", file=sys.stderr)


def note(message: str) -> None:
    """``note: <message>``: worth reading, but nothing failed."""
    print(f"{paint('note:', 'bold', 'yellow')} {message}")


def detail(message: str) -> None:
    """A bookkeeping line, greyed so the results stand out."""
    print(paint(message, "grey"))


def made(what: str, where: Path | str, after: str = "") -> None:
    """``<what> -> <where><after>``: a file the run wrote."""
    print(f"{paint(what + ' ->', 'bold', 'green')} {path(where)}{after}")
