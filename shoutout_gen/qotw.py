"""Tally the form's question-of-the-week column and draw it as a bar chart.

Every week the form carries a QOTW whose answer is a pick from a list of names;
the question itself lives in the column header ("QOTW [9/30]: Which board member
is the greenest flag?") and is edited each week, while older answers stay in the
same column. So, like the shout-outs, a run needs ``--since`` to keep only the
current week's answers.

The column is found by its header (``QOTW`` / ``question of the week``), falling
back to column D where the form put it. A checkbox question exports several
picks as one cell joined by ", "; each pick is counted once. Picks are merged
case- and spacing-insensitively and shown in their most common spelling.

The chart is a 16:9 PNG -- names ordered most picked to least, count over each
bar -- so it can go full-bleed on a slide or be pasted into one by hand.
"""
from __future__ import annotations

import re
import string
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .deck import HAUSCR_RED
from .sheet import TIME_HEADER_RE, SheetFormatError, parse_timestamp, read_table

# Anchored: a shout-out header that only mentions the QOTW ("... then answer the QOTW!") is not it.
QOTW_HEADER_RE = re.compile(r"^\s*(?:qotw|question of the week)", re.I)
FALLBACK_COLUMN = "D"  # where the form has put the QOTW so far
# "QOTW [9/30]: Which ...?" -> "Which ...?"; the tag in front is bookkeeping, not the question.
_QUESTION_PREFIX_RE = re.compile(r"^\s*(?:qotw|question of the week)\s*(?:\[[^\]]*\]|\([^)]*\))?\s*[:\-–—]?\s*", re.I)
_PICK_SEPARATOR = ", "  # how Google Forms joins checkbox answers in one cell
_COLUMN_LETTER_RE = re.compile(r"[A-Za-z]{1,2}")

# Slide look: the club template's white page, Roboto, theme ink, and the bars in
# the club's red, the same red as the "QOTW:" pane before the chart.
BAR_COLOUR = f"#{HAUSCR_RED}"
INK = "#000000"
MUTED_INK = "#595959"  # theme dk2
AXIS_COLOUR = "#BDBDBD"
FIG_SIZE_IN = (13.333, 7.5)  # 16:9, the template's page shape
DPI = 150
TITLE_PT = (28, 26, 24, 22)  # the question shrinks through these to stay on one line, then wraps at the last
TITLE_WIDTH = 0.92  # of the figure: the same width as the bars


@dataclass(frozen=True)
class Tally:
    """Picks for one question, most picked first (ties alphabetical)."""

    question: str
    counts: list[tuple[str, int]]
    responses: int
    first: datetime | None
    last: datetime | None


def output_name_for(gm_number: int, date: str, suffix: str) -> str:
    """Beside the shout-out deck: "GM #26 10.7---qotw.png" / ".pptx"."""
    return f"GM #{gm_number} {date}---qotw{suffix}"


def read_tally(path: Path, column: str | None = None, since: datetime | None = None) -> tuple[Tally, int]:
    """Count the QOTW picks in ``path``; returns the tally and how many answers ``since`` left out.

    ``column`` forces the column, by header name or by letter ("D").
    """
    header, rows = read_table(Path(path))
    if not header:
        raise SheetFormatError(f"{path} is empty")
    col = pick_column(header, column)
    time_col = next((h for h in header if TIME_HEADER_RE.search(h)), None)

    answers: list[tuple[str, datetime | None]] = []
    skipped = 0
    for row in rows:
        cell = _clean(row.get(col))
        if not cell:
            continue
        stamp = parse_timestamp(row.get(time_col)) if time_col else None
        if since is not None and (stamp is None or stamp < since):
            skipped += 1
            continue
        answers.append((cell, stamp))
    return tally(question_text(col), answers), skipped


def pick_column(header: list[str], forced: str | None = None) -> str:
    """The QOTW column: ``forced`` (header name or letter), else by header, else column D."""
    if forced is not None:
        if forced in header:
            return forced
        if _COLUMN_LETTER_RE.fullmatch(forced):
            index = _letter_index(forced)
            if index < len(header):
                return header[index]
        raise SheetFormatError(f"column {forced!r} is neither a header nor a column letter of {header}")
    for name in header:
        if QOTW_HEADER_RE.search(name):
            return name
    index = _letter_index(FALLBACK_COLUMN)
    if index < len(header) and header[index]:
        return header[index]
    raise SheetFormatError(f"no QOTW column in header {header}; pass --column")


def question_text(header_cell: str) -> str:
    """The question as asked: the header minus its "QOTW [date]:" tag."""
    stripped = _QUESTION_PREFIX_RE.sub("", header_cell).strip()
    return stripped or header_cell.strip()


def tally(question: str, answers: list[tuple[str, datetime | None]]) -> Tally:
    """Count every pick in ``answers`` (cell text, timestamp)."""
    answers = [(cell, stamp) for cell, stamp in answers if _picks(cell)]  # a lone "/" is no response
    counts: Counter[str] = Counter()
    spellings: dict[str, Counter[str]] = defaultdict(Counter)
    for cell, _ in answers:
        for pick in {_key(p): p for p in _picks(cell)}.values():  # one vote per name per response
            counts[_key(pick)] += 1
            spellings[_key(pick)][pick] += 1
    shown = {key: spellings[key].most_common(1)[0][0] for key in counts}
    ordered = sorted(counts.items(), key=lambda kv: (-kv[1], shown[kv[0]].casefold()))
    stamps = [s for _, s in answers if s is not None]
    return Tally(
        question=question,
        counts=[(shown[key], n) for key, n in ordered],
        responses=len(answers),
        first=min(stamps, default=None),
        last=max(stamps, default=None),
    )


def subtitle_for(result: Tally) -> str:
    """"Question of the week · 27 responses · 10/1 – 10/7" (dates as the form's headers write them)."""
    noun = "response" if result.responses == 1 else "responses"
    span = ""
    if result.first and result.last:
        span = f" · {result.first.month}/{result.first.day} – {result.last.month}/{result.last.day}"
    return f"Question of the week · {result.responses} {noun}{span}"


def render_chart(result: Tally, out_path: Path, font_file: Path | None = None) -> Path:
    """Draw ``result`` as a 16:9 bar chart PNG at ``out_path``; returns it."""
    subtitle = subtitle_for(result)
    import matplotlib

    matplotlib.use("Agg")
    from matplotlib import font_manager
    from matplotlib import pyplot as plt

    family = "sans-serif"
    if font_file is not None:
        font_manager.fontManager.addfont(str(font_file))
        family = font_manager.FontProperties(fname=str(font_file)).get_name()

    names = [name for name, _ in result.counts]
    values = [n for _, n in result.counts]
    # Names go under their bar on two lines ("Katherine\nCai"); shrink the type once the row gets crowded.
    label_pt = 15 if len(names) <= 12 else max(8.0, 15 * 12 / len(names))

    with plt.rc_context({"font.family": family, "text.color": INK}):
        fig = plt.figure(figsize=FIG_SIZE_IN, dpi=DPI, facecolor="white")
        title = fig.text(0.04, 0.93, result.question, fontsize=TITLE_PT[0], va="top", ha="left")
        one_line = _fit_title(title, fig)
        drop = _height(title, fig) - one_line  # a wrapped question pushes the rest down
        if subtitle:
            fig.text(0.04, 0.855 - drop, subtitle, fontsize=15, color=MUTED_INK, va="top", ha="left")
        ax = fig.add_axes((0.04, 0.17, 0.92, 0.62 - drop))
        positions = range(len(names))
        ax.bar(positions, values, width=0.68, color=BAR_COLOUR, zorder=2)
        for x, n in zip(positions, values):
            ax.annotate(str(n), (x, n), xytext=(0, 6), textcoords="offset points",
                        ha="center", va="bottom", fontsize=label_pt + 3)
        ax.set_xticks(list(positions), [name.replace(" ", "\n", 1) for name in names], fontsize=label_pt)
        ax.set_xlim(-0.6, max(len(names), 1) - 0.4)
        ax.set_ylim(0, max(values, default=1) * 1.15)
        ax.tick_params(axis="x", length=0, pad=8)
        ax.yaxis.set_visible(False)
        for side in ("top", "right", "left"):
            ax.spines[side].set_visible(False)
        ax.spines["bottom"].set_color(AXIS_COLOUR)
        ax.set_facecolor("white")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        fig.savefig(out_path, dpi=DPI, facecolor="white")
        plt.close(fig)
    return out_path


# --------------------------------------------------------------------------- #
# Private mechanics
# --------------------------------------------------------------------------- #
_NOT_A_PICK = set(string.punctuation + string.whitespace)


def _width(text, fig) -> float:
    return text.get_window_extent(fig.canvas.get_renderer()).width / fig.bbox.width


def _height(text, fig) -> float:
    return text.get_window_extent(fig.canvas.get_renderer()).height / fig.bbox.height


def _fit_title(title, fig) -> float:
    """Shrink ``title`` until it fits TITLE_WIDTH, else wrap it at the smallest size.

    Returns the height of one line at the starting size, the room the layout below assumes.
    """
    one_line = _height(title, fig)
    for size in TITLE_PT:
        title.set_fontsize(size)
        if _width(title, fig) <= TITLE_WIDTH:
            return one_line
    lines: list[str] = []
    for word in title.get_text().split():
        title.set_text(f"{lines[-1]} {word}" if lines else word)
        if not lines or _width(title, fig) <= TITLE_WIDTH:
            lines[-1:] = [title.get_text()]
        else:
            lines.append(word)
    title.set_text("\n".join(lines))
    return one_line


def _picks(cell: str) -> list[str]:
    parts = (" ".join(p.split()) for p in cell.split(_PICK_SEPARATOR))
    return [p for p in parts if p and not set(p) <= _NOT_A_PICK]


def _key(pick: str) -> str:
    return " ".join(pick.split()).casefold()


def _clean(value) -> str:
    return "" if value is None else str(value).strip()


def _letter_index(letters: str) -> int:
    index = 0
    for ch in letters.upper():
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index - 1
