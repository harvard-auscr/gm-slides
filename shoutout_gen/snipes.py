"""The snipe game's GM slide: this week's leaderboard, then the season so far.

Reads the day-by-day workbook that ``snipebot export`` writes
(``<semester>_day_by_day.xlsx``). Its Snipes sheet has one row per counted snipe
-- a sniper-target pair, so a photo tagging three people is three rows -- with
the points each side scored, cooldowns, vetoes and selfie bonuses already
applied. Nothing here re-judges a snipe; it only counts rows. People are told
apart by the sheet's person keys, since two members can share a display name.

Two 16:9 PNGs come out:

  week    "Most snipes made" and "Most sniped" in the window (counts), the same
          boards as snipebot's tools/leaderboard.py
  season  the race for the lead: the season's top scorers' running points at the
          end of each week, next to the season's points and most-sniped lists
"""
from __future__ import annotations

import math
import re
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Hashable, Mapping

DEFAULT_EXPORTS = Path(__file__).resolve().parents[2] / "snipe-tracker" / "exports"
WORKBOOK_GLOB = "*_day_by_day.xlsx"
SNIPES_SHEET = "Snipes"
TOP = 10
TIE_SLACK = 5  # show everyone tied at the cutoff when that adds at most this many rows
RACE_LINES = 5

_COLUMNS = {"date": "Date", "time": "Time", "sniper": "Sniper", "target": "Target",
            "sniper_key": "Sniper key", "target_key": "Target key",
            "sniper_points": "Sniper points", "target_points": "Target points"}
# The export prefixes a name with ' when it would otherwise read as a spreadsheet formula.
_GUARDED_NAME_RE = re.compile(r"'([=+\-@\t\r].*)", re.S)


class SnipeExportError(ValueError):
    """The file is not a readable snipebot day-by-day workbook."""


@dataclass(frozen=True)
class Snipe:
    at: datetime  # club-local time, as the workbook prints it
    sniper: str
    target: str
    sniper_points: int
    target_points: int
    sniper_id: Hashable  # the person key; falls back to the name when the sheet has none
    target_id: Hashable


@dataclass(frozen=True)
class Entry:
    rank: int
    name: str
    value: int
    id: Hashable


@dataclass(frozen=True)
class Board:
    entries: list[Entry]
    hidden: int       # people left off below the cutoff
    hidden_best: int  # the best value among those left off (0 if none)

    def leaders(self) -> list[Entry]:
        return [e for e in self.entries if e.rank == 1]


@dataclass(frozen=True)
class Week:
    since: datetime
    until: datetime
    made: Board
    sniped: Board
    snipes: int
    snipers: int
    people_sniped: int
    busiest_day: tuple[date, int] | None


@dataclass(frozen=True)
class RaceLine:
    name: str
    id: Hashable
    values: list[int]  # running points at each week end


@dataclass(frozen=True)
class Season:
    first: date
    until: datetime
    snipes: int
    players: int
    points: Board
    sniped: Board
    week_ends: list[date]
    race: list[RaceLine]  # best first


# --------------------------------------------------------------------------- #
# Reading
# --------------------------------------------------------------------------- #
def find_workbook(path: Path) -> Path:
    """``path`` itself, or the newest day-by-day workbook in that folder."""
    if path.is_file():
        return path
    if not path.is_dir():
        raise FileNotFoundError(f"no such file or folder: {path}")
    found = [p for p in path.glob(WORKBOOK_GLOB) if not p.name.startswith("~$")]
    if not found:
        raise FileNotFoundError(f"no {WORKBOOK_GLOB} in {path}; run snipebot export first")
    return max(found, key=lambda p: (p.stat().st_mtime, p.name))


def load_snipes(path: Path) -> list[Snipe]:
    """Every row of the workbook's Snipes sheet, in sheet order."""
    from openpyxl import load_workbook
    from openpyxl.utils.exceptions import InvalidFileException

    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except (InvalidFileException, zipfile.BadZipFile, KeyError, OSError, ValueError) as exc:
        raise SnipeExportError(f"cannot read {path.name} as a workbook ({exc}); pass snipebot's *_day_by_day.xlsx") from exc
    try:
        if SNIPES_SHEET not in wb.sheetnames:
            raise SnipeExportError(f"{path.name} has no {SNIPES_SHEET!r} sheet; is it a snipebot day-by-day workbook?")
        rows = wb[SNIPES_SHEET].iter_rows(values_only=True)
        header = [str(c).strip() if c is not None else "" for c in next(rows, ())]
        missing = [h for h in _COLUMNS.values() if h not in header]
        if missing:
            raise SnipeExportError(f"{path.name}'s {SNIPES_SHEET} sheet is missing {missing}")
        col = {key: header.index(h) for key, h in _COLUMNS.items()}
        out = []
        for number, row in enumerate(rows, 2):
            row = tuple(row) + (None,) * (len(header) - len(row))
            sniper, target = _name(row[col["sniper"]]), _name(row[col["target"]])
            if row[col["date"]] is None or not sniper or not target:
                continue
            try:
                out.append(Snipe(
                    _when(row[col["date"]], row[col["time"]]), sniper, target,
                    _int(row[col["sniper_points"]]), _int(row[col["target_points"]]),
                    _id(row[col["sniper_key"]], sniper), _id(row[col["target_key"]], target),
                ))
            except (TypeError, ValueError) as exc:
                raise SnipeExportError(f"{path.name} {SNIPES_SHEET} row {number}: {exc}") from exc
        return out
    finally:
        wb.close()


def _when(day, clock) -> datetime:
    d = day.date() if isinstance(day, datetime) else day if isinstance(day, date) else date.fromisoformat(str(day)[:10])
    if isinstance(clock, time):
        return datetime.combine(d, clock)
    if isinstance(clock, datetime):
        return datetime.combine(d, clock.time())
    return datetime.combine(d, time.fromisoformat(str(clock).strip())) if clock else datetime.combine(d, time())


def _name(cell) -> str:
    text = str(cell or "").strip()
    guarded = _GUARDED_NAME_RE.fullmatch(text)
    return guarded.group(1) if guarded else text


def _int(cell) -> int:
    value = float(cell or 0)
    if not value.is_integer():
        raise ValueError(f"points {cell!r} is not a whole number")
    return int(value)


def _id(key, name: str) -> Hashable:
    if key is None or str(key).strip() == "":
        return name
    return int(key) if isinstance(key, (int, float)) and float(key).is_integer() else str(key)


# --------------------------------------------------------------------------- #
# Counting
# --------------------------------------------------------------------------- #
def rank(values: Counter, names: Mapping[Hashable, str] | None = None, top: int = TOP) -> Board:
    """Most first, then by name; standard competition ranks (1, 2, 2, 4). Cut at ``top``,
    but keep everyone tied at the cutoff if that adds no more than TIE_SLACK rows.
    ``values`` is keyed by person id; ``names`` maps ids to display names (default: the id)."""
    def name(pid) -> str:
        return names.get(pid, str(pid)) if names is not None else str(pid)

    ordered = sorted(((pid, v) for pid, v in values.items() if v > 0),
                     key=lambda kv: (-kv[1], name(kv[0]).casefold(), name(kv[0]), str(kv[0])))
    entries, prev, rank_no = [], None, 0
    for i, (pid, v) in enumerate(ordered, 1):
        if v != prev:
            rank_no, prev = i, v
        entries.append(Entry(rank_no, name(pid), v, pid))
    cut = min(top, len(entries))
    if 0 < cut < len(entries):
        tied = cut
        while tied < len(entries) and entries[tied].value == entries[cut - 1].value:
            tied += 1
        if tied - cut <= TIE_SLACK:
            cut = tied
    rest = entries[cut:]
    return Board(entries[:cut], len(rest), rest[0].value if rest else 0)


def end_of_day(d: date) -> datetime:
    return datetime.combine(d, time.max)


def week(snipes: list[Snipe], since: datetime, until: datetime, top: int = TOP) -> Week:
    """Counts in ``since <= at <= until``: one row is one snipe made and one snipe taken."""
    rows = [s for s in snipes if since <= s.at <= until]
    names = _names(rows)
    by_day = Counter(s.at.date() for s in rows)
    busiest = min(by_day.items(), key=lambda kv: (-kv[1], kv[0])) if by_day else None
    return Week(
        since, until,
        rank(Counter(s.sniper_id for s in rows), names, top), rank(Counter(s.target_id for s in rows), names, top),
        len(rows), len({s.sniper_id for s in rows}), len({s.target_id for s in rows}), busiest,
    )


def season(snipes: list[Snipe], until: datetime, top: int = TOP, race_lines: int = RACE_LINES) -> Season:
    """Everything up to ``until``: points and most-sniped boards, and the leaders' race
    as running points at the end of each 7-day week counted back from ``until``. The
    race takes the top ``race_lines`` by points, plus anyone tied at the cut, up to
    one line per race colour."""
    rows = [s for s in snipes if s.at <= until]
    if not rows:
        raise ValueError(f"no snipes on or before {until:%Y-%m-%d %H:%M}")
    names = _names(rows)
    points = _points(rows)
    first = min(s.at for s in rows).date()
    week_ends = []
    end = until.date()
    while end >= first:
        week_ends.append(end)
        end -= timedelta(days=7)
    week_ends.reverse()

    totals = [_points([s for s in rows if s.at <= (until if d == until.date() else end_of_day(d))])
              for d in week_ends]
    race = [RaceLine(e.name, e.id, [t[e.id] for t in totals])
            for e in rank(points, names, race_lines).entries[: len(RACE_COLOURS)]]
    return Season(
        first, until, len(rows), len({s.sniper_id for s in rows} | {s.target_id for s in rows}),
        rank(points, names, top), rank(Counter(s.target_id for s in rows), names, top), week_ends, race,
    )


def _points(rows: list[Snipe]) -> Counter:
    points: Counter = Counter()
    for s in rows:
        points[s.sniper_id] += s.sniper_points
        points[s.target_id] += s.target_points
    return points


def _names(rows: list[Snipe]) -> dict[Hashable, str]:
    names: dict[Hashable, str] = {}
    for s in rows:
        names[s.sniper_id] = s.sniper
        names[s.target_id] = s.target
    return names


def leader_text(board: Board, unit: str) -> str:
    """"Benjamin Mujkic (280 points)", "A and B (48 points each)" or "6 tied (48 points)"."""
    top = board.leaders()
    if not top:
        return ""
    n = top[0].value
    if len(top) == 1:
        return f"{top[0].name} ({n} {unit})"
    if len(top) == 2:
        return f"{top[0].name} and {top[1].name} ({n} {unit} each)"
    return f"{len(top)} tied ({n} {unit} each)"


def plural(n: int, word: str) -> str:
    return f"{n:,} {word}" + ("" if n == 1 else "s")


# --------------------------------------------------------------------------- #
# Drawing
# --------------------------------------------------------------------------- #
# White page like the club template; ink and grid one step off it. Blue/orange are
# the week boards' colours (snipebot's leaderboard PNG); the race takes the
# categorical slots in order, one per leader, each line labelled with its name.
SURFACE, INK, INK_2, GRID = "#ffffff", "#0b0b0b", "#52514e", "#e4e3df"
MADE_COLOUR, SNIPED_COLOUR = "#2a78d6", "#eb6834"
RACE_COLOURS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
# Names in the club include CJK; Roboto has none, so the fallbacks carry them.
FALLBACK_FONTS = ["Microsoft YaHei", "Malgun Gothic", "Yu Gothic", "Noto Sans CJK SC", "DejaVu Sans"]
FIG_SIZE_IN = (13.333, 7.5)  # 16:9, the template's page shape
DPI = 150
MAX_WEEK_TICKS = 9  # more weeks than this and only every k-th week end is labelled

# Week boards: left edge of each board's bars, its width, and the room its names get.
WEEK_BOARDS_X = (0.19, 0.68)
WEEK_BOARD_W = 0.27
WEEK_LABEL_IN = 1.95
# Season: the race chart's box, the room for its end labels, and the two lists.
RACE_BOX = (0.04, 0.10, 0.34, 0.64)
RACE_LABEL_IN = 2.4
LISTS = ((0.59, "Season points"), (0.80, "Most sniped"))
LIST_W = 0.18
LIST_TOP, LIST_BOTTOM = 0.735, 0.10  # figure fractions the ranked lists' rows span
LIST_TITLE_Y = 0.762  # the race chart's title baseline, so the three titles line up


def window_label(since: datetime, until: datetime) -> str:
    def fmt(t: datetime, end: bool) -> str:
        day = f"{t:%a} {t:%b} {t.day}"
        if since.year != until.year:
            day += f", {t.year}"
        bare = t.time() >= time(23, 59, 59) if end else t.time() == time()
        return day if bare else f"{day}, {t:%H:%M}"
    return f"{fmt(since, False)} – {fmt(until, True)}"


def render_week(w: Week, out_path: Path, font_file: Path | None = None) -> Path:
    plt = _pyplot(font_file)
    fig = plt.figure(figsize=FIG_SIZE_IN, dpi=DPI, facecolor=SURFACE)
    stats = f"{plural(w.snipes, 'snipe')} by {plural(w.snipers, 'sniper')} · {w.people_sniped} people sniped"
    if w.busiest_day:
        d, n = w.busiest_day
        stats += f" · busiest day {d:%a %b} {d.day} ({n})"
    _header(fig, f"Snipe leaderboard · {window_label(w.since, w.until)}", stats)
    rows = max(len(w.made.entries), len(w.sniped.entries), 1)
    for x, title, board, colour in ((WEEK_BOARDS_X[0], "Most snipes made", w.made, MADE_COLOUR),
                                    (WEEK_BOARDS_X[1], "Most sniped", w.sniped, SNIPED_COLOUR)):
        _board(fig, fig.add_axes((x, 0.09, WEEK_BOARD_W, 0.67)), title, board, colour, rows,
               "no snipes in this window")
    return _save(plt, fig, out_path)


def render_season(s: Season, out_path: Path, font_file: Path | None = None) -> Path:
    plt = _pyplot(font_file)
    fig = plt.figure(figsize=FIG_SIZE_IN, dpi=DPI, facecolor=SURFACE)
    stats = f"{plural(s.snipes, 'snipe')} · {plural(s.players, 'player')} · {plural(len(s.week_ends), 'week')}"
    lead = leader_text(s.points, "points")
    if lead:
        stats += f" · in the lead: {lead}"
    _header(fig, f"Season so far · {window_label(datetime.combine(s.first, time()), s.until)}", stats)

    colour_of = {line.id: colour for line, colour in zip(s.race, RACE_COLOURS)}
    _race(fig, fig.add_axes(RACE_BOX), s)
    rows = max(len(s.points.entries), len(s.sniped.entries), 1)
    (x1, t1), (x2, t2) = LISTS
    _ranked_list(fig, x1, LIST_W, t1, s.points, rows, "no points yet", colour_of)
    _ranked_list(fig, x2, LIST_W, t2, s.sniped, rows, "nobody sniped yet")
    return _save(plt, fig, out_path)


def fit_text(fig, text: str, fontsize: float, max_width_in: float) -> str:
    """``text``, cut with an ellipsis until it renders no wider than ``max_width_in``."""
    renderer = fig.canvas.get_renderer()
    probe = fig.text(0, 0, text, fontsize=fontsize)
    try:
        def width(s: str) -> float:
            probe.set_text(s)
            return probe.get_window_extent(renderer).width / fig.dpi

        if width(text) <= max_width_in:
            return text
        cut = len(text) - 1
        while cut > 1 and width(text[:cut].rstrip() + "…") > max_width_in:
            cut -= 1
        return text[:cut].rstrip() + "…"
    finally:
        probe.remove()


def text_width_in(fig, text: str, fontsize: float) -> float:
    renderer = fig.canvas.get_renderer()
    probe = fig.text(0, 0, text, fontsize=fontsize)
    try:
        return probe.get_window_extent(renderer).width / fig.dpi
    finally:
        probe.remove()


def label_positions(values: list[float], gap: float, floor: float = 0.0) -> list[float]:
    """Heights for end labels in best-first order: each at least ``gap`` below the one
    above, none below ``floor`` (the stack is pushed back up from the floor instead)."""
    placed: list[float] = []
    for v in values:
        placed.append(min(v, placed[-1] - gap) if placed else v)
    if placed and placed[-1] < floor:
        placed[-1] = floor
        for i in range(len(placed) - 2, -1, -1):
            placed[i] = max(placed[i], placed[i + 1] + gap)
    return placed


def week_ticks(week_ends: list[date]) -> list[tuple[int, str]]:
    """(index, label) for the race's x axis: every week up to MAX_WEEK_TICKS, else every
    k-th one counted back from the last, so the newest week is always labelled."""
    k = max(1, math.ceil(len(week_ends) / MAX_WEEK_TICKS))
    last = len(week_ends) - 1
    return [(i, f"{d.month}/{d.day}") for i, d in enumerate(week_ends) if (last - i) % k == 0]


def _pyplot(font_file: Path | None):
    import matplotlib

    matplotlib.use("Agg")
    from matplotlib import font_manager
    from matplotlib import pyplot as plt

    families = []
    if font_file is not None:
        font_manager.fontManager.addfont(str(font_file))
        families.append(font_manager.FontProperties(fname=str(font_file)).get_name())
    have = {f.name for f in font_manager.fontManager.ttflist}
    families += [f for f in FALLBACK_FONTS if f in have]
    plt.rcParams["font.family"] = families or ["sans-serif"]
    plt.rcParams["svg.hashsalt"] = "gm-slides"
    return plt


def _header(fig, title: str, stats: str) -> None:
    fig.text(0.04, 0.93, title, fontsize=26, color=INK, va="top")
    fig.text(0.04, 0.855, stats, fontsize=15, color=INK_2, va="top")


def _board(fig, ax, title: str, board: Board, colour: str, rows: int, empty: str) -> None:
    ax.set_facecolor(SURFACE)
    entries = board.entries
    ys = list(range(len(entries)))
    ax.barh(ys, [e.value for e in entries], height=0.5, color=colour)
    labels = [fit_text(fig, f"{e.rank}.  {e.name}", 14, WEEK_LABEL_IN) for e in entries]
    ax.set_yticks(ys, labels, fontsize=14, color=INK)
    ax.set_ylim(rows - 0.5, -0.5)
    top = max((e.value for e in entries), default=1)
    ax.set_xlim(0, top * 1.16)
    for y, e in zip(ys, entries):
        ax.text(e.value + top * 0.02, y, str(e.value), va="center", fontsize=13, color=INK)
    ax.set_title(title, loc="left", fontsize=17, color=INK, pad=12)
    if board.hidden:
        ax.set_xlabel(f"+{board.hidden} more with {board.hidden_best} or fewer", loc="left",
                      fontsize=12, color=INK_2)
    if not entries:
        ax.text(0.5, 0.5, empty, transform=ax.transAxes, ha="center", color=INK_2, fontsize=14)
    ax.tick_params(axis="y", length=0, pad=8)
    ax.tick_params(axis="x", colors=INK_2, labelsize=11)
    ax.xaxis.get_major_locator().set_params(integer=True)
    ax.grid(axis="x", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(GRID)


def _ranked_list(fig, x: float, width: float, title: str, board: Board, rows: int, empty: str,
                 marks: Mapping[Hashable, str] | None = None) -> None:
    """A season board as text: rank, name, value right-aligned; a leader's race colour as a dot."""
    fig.text(x, LIST_TITLE_Y, title, fontsize=17, color=INK, va="baseline")
    step = (LIST_TOP - LIST_BOTTOM) / max(rows, 1)
    w_in = fig.get_size_inches()[0]
    name_x = x + 0.032
    for i, e in enumerate(board.entries):
        y = LIST_TOP - (i + 0.5) * step
        value = str(e.value)
        room = (x + width - name_x) * w_in - text_width_in(fig, value, 14) - 0.15
        fig.text(x + 0.022, y, f"{e.rank}.", fontsize=14, color=INK_2, va="center", ha="right")
        fig.text(name_x, y, fit_text(fig, e.name, 14, room), fontsize=14, color=INK, va="center")
        fig.text(x + width, y, value, fontsize=14, color=INK, va="center", ha="right")
        if marks and e.id in marks:
            fig.patches.append(_dot(fig, x - 0.008, y, marks[e.id]))
        fig.add_artist(_rule(fig, x, x + width, y - step / 2))
    if board.hidden:
        fig.text(x, LIST_BOTTOM - 0.035, f"+{board.hidden} more with {board.hidden_best} or fewer",
                 fontsize=12, color=INK_2, va="top")
    if not board.entries:
        fig.text(x, LIST_TOP - step, empty, fontsize=14, color=INK_2, va="center")


def _dot(fig, x: float, y: float, colour: str):
    from matplotlib.patches import Ellipse

    w_in, h_in = fig.get_size_inches()
    d = 0.11  # inches
    return Ellipse((x, y), d / w_in, d / h_in, transform=fig.transFigure, color=colour)


def _rule(fig, x0: float, x1: float, y: float):
    from matplotlib.lines import Line2D

    return Line2D([x0, x1], [y, y], transform=fig.transFigure, color=GRID, linewidth=0.8)


def _race(fig, ax, s: Season) -> None:
    ax.set_facecolor(SURFACE)
    xs = list(range(len(s.week_ends)))
    top = max((max(line.values) for line in s.race), default=1) or 1
    for line, colour in zip(s.race, RACE_COLOURS):
        ax.plot(xs, line.values, color=colour, linewidth=3, solid_joinstyle="round", solid_capstyle="round", zorder=3)
        ax.plot(xs[-1:], line.values[-1:], marker="o", markersize=9, color=colour,
                markeredgecolor=SURFACE, markeredgewidth=2, zorder=4)
    # End labels, nudged apart where the leaders finish close together, never below zero.
    gap = top * 0.075
    finals = [line.values[-1] for line in s.race]
    placed = label_positions(finals, gap)
    for line, final, y in zip(s.race, finals, placed):
        value = f"  {final}"
        name = fit_text(fig, line.name, 13, RACE_LABEL_IN - text_width_in(fig, value, 13))
        ax.annotate(name + value, (xs[-1], final), xytext=(xs[-1] + 0.25, y),
                    textcoords="data", va="center", fontsize=13, color=INK, annotation_clip=False,
                    arrowprops=dict(arrowstyle="-", color=INK_2, linewidth=0.8, shrinkA=2, shrinkB=6,
                                    relpos=(0, 0.5)) if y != final else None)
    ticks = week_ticks(s.week_ends)
    ax.set_xticks([i for i, _ in ticks], [label for _, label in ticks], fontsize=11, color=INK_2)
    ax.set_xlim(-0.3, max(len(xs) - 1, 1) + 0.3)
    ax.set_ylim(0, max(top * 1.08, max(placed, default=0) + gap / 2))
    ax.set_title("Race for the lead: points by week", loc="left", fontsize=17,
                 color=INK, pad=12)
    ax.set_xlabel("week ending", loc="left", fontsize=12, color=INK_2)
    ax.tick_params(axis="both", length=0, colors=INK_2, labelsize=11)
    ax.yaxis.get_major_locator().set_params(integer=True)
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(GRID)


def _save(plt, fig, out_path: Path) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fig.savefig(out_path, dpi=DPI, facecolor=SURFACE, metadata={"Software": None})
    plt.close(fig)
    return out_path


def output_name_for(gm_number: int, date_text: str, suffix: str) -> str:
    """"GM #7 10.28---snipes.pptx", beside the shout-out deck's name."""
    return f"GM #{gm_number} {date_text}---snipes{suffix}"
