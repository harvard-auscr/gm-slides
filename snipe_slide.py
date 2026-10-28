"""Make the snipe game's GM slide: this week's leaderboard, then the season on a click.

    python snipe_slide.py --gm 7 --date 10.28 --since 10.22
        -> reads the newest *_day_by_day.xlsx in ../snipe-tracker/exports and writes
           output/GM #7 10.28---snipes.pptx (+ ---snipes-week.png, ---snipes-season.png)
    python snipe_slide.py path/to/fall-2026_day_by_day.xlsx --gm 7 --date 10.28 --since 10.22

The workbook is the one `snipebot export` writes (also in the weekly `standings`
artifact); recount and export first so the slide has the latest snipes. The slide
opens on this week's boards -- most snipes made and most sniped -- and one click
puts the season over it: the race for the lead (the top scorers' running points
week by week) beside the season's points and most-sniped lists.

--since and --until take m.d (e.g. 10.22), 'm.d HH:MM' or YYYY-MM-DD [HH:MM]; an
m.d with no year is the most recent such day that is not in the future (one more
than half a year back is refused as a typo: give the full date). The week
runs from --since (a bare day = its start) through --until (a bare day = its end;
default: the end of today). Without --since it is the 7 days ending at --until.
The season runs from the first snipe through --until.

If ``.env`` names a Google Drive folder the slide is uploaded there as Google
Slides, ready for File -> Import slides in the GM deck.
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import date, datetime, time, timedelta
from pathlib import Path

from shoutout_gen import cli, drive, term
from shoutout_gen.deck import TEMPLATE_PATH, write_reveal_slide
from shoutout_gen.metrics import FontNotFoundError, resolve_font
from shoutout_gen.sheet import parse_since
from shoutout_gen.snipes import (DEFAULT_EXPORTS, SnipeExportError, end_of_day, find_workbook, leader_text,
                                 load_snipes, output_name_for, plural, render_season, render_week, season, week,
                                 window_label)

_HAS_TIME_RE = re.compile(r"\d{1,2}:\d{2}|T\d", re.I)


def parse_moment(text: str, today: date, end: bool) -> datetime:
    """--since / --until: a bare day is its start, or for --until (``end``) its end."""
    moment = parse_since(text, today)
    if end and moment.time() == time() and not _HAS_TIME_RE.search(text):
        return end_of_day(moment.date())
    return moment


def _main(argv: list[str] | None = None, today: date | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):  # a name the console can't encode prints as "?" instead of crashing
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")
    today = today or date.today()
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument(
        "export", nargs="?", type=Path, default=DEFAULT_EXPORTS,
        help=f"snipebot's day-by-day workbook, or a folder holding one (default: the newest in {DEFAULT_EXPORTS})",
    )
    parser.add_argument("--gm", type=int, required=True, help="general meeting number, e.g. 7")
    parser.add_argument("--date", required=True, help="date as written in deck names, e.g. 10.28")
    parser.add_argument("--since", help="start of the week (default: 7 days before --until)")
    parser.add_argument("--until", help="end of the week and of the season (default: the end of today)")
    parser.add_argument("--out", type=Path, default=Path("output"), help="output folder")
    parser.add_argument("--template", type=Path, default=TEMPLATE_PATH)
    parser.add_argument("--font", default=None, help="typeface for the charts (default Roboto)")
    parser.add_argument("--font-file", type=Path, default=None, help="TTF/OTF to draw with (default: the bundled Roboto)")
    parser.add_argument("--env-file", type=Path, default=cli.DEFAULT_ENV_FILE, help="Google settings (see .env.example)")
    cli.add_upload_args(parser)
    args = parser.parse_args(argv)

    try:
        end = parse_moment(args.until, today, end=True) if args.until else end_of_day(today)
        start = parse_moment(args.since, today, end=False) if args.since else datetime.combine(end.date() - timedelta(days=6), time())
    except ValueError as exc:
        term.error(str(exc))
        return 2
    if start > end:
        term.error(f"--since ({start:%Y-%m-%d %H:%M}) is after --until ({end:%Y-%m-%d %H:%M})")
        return 2

    try:
        upload_cfg = None if args.no_upload else drive.load_config(args.env_file)
    except drive.DriveConfigError as exc:
        term.error(str(exc))
        return 2
    try:
        _, font_file = resolve_font(args.font, args.font_file)
    except FontNotFoundError as exc:
        term.error(str(exc))
        return 2
    try:
        workbook = find_workbook(args.export)
        snipes = load_snipes(workbook)
        this_week = week(snipes, start, end)
        so_far = season(snipes, end)
    except (FileNotFoundError, SnipeExportError, ValueError) as exc:
        term.error(str(exc))
        return 2
    term.detail(f"read {len(snipes)} snipes from {workbook}")
    if not this_week.snipes:
        term.note(f"no snipes {window_label(start, end)}; the week's boards will be empty")

    week_png = render_week(this_week, args.out / output_name_for(args.gm, args.date, "-week.png"), font_file)
    season_png = render_season(so_far, args.out / output_name_for(args.gm, args.date, "-season.png"), font_file)
    pptx = write_reveal_slide(week_png, season_png, args.out / output_name_for(args.gm, args.date, ".pptx"),
                              args.template, names=("Snipes this week", "Snipes this season"))

    print()
    print(term.paint(f"This week ({window_label(start, end)})", "bold"))
    print(f"  {plural(this_week.snipes, 'snipe')} by {plural(this_week.snipers, 'sniper')}; "
          f"{this_week.people_sniped} people sniped")
    for label, board in (("most snipes made", this_week.made), ("most sniped", this_week.sniped)):
        top3 = ", ".join(f"{e.name} {e.value}" for e in board.entries[:3])
        print(f"  {label}: {top3 or '-'}")
    print(term.paint("Season", "bold"))
    print(f"  {plural(so_far.snipes, 'snipe')}, {plural(so_far.players, 'player')}, {plural(len(so_far.week_ends), 'week')}")
    print(f"  in the lead: {leader_text(so_far.points, 'points')}")
    print("  race: " + ", ".join(f"{line.name} {line.values[-1]}" for line in so_far.race))
    print()
    term.made("week", week_png)
    term.made("season", season_png)
    term.made("slide", pptx, " (1 click: the season appears over the week)")

    if upload_cfg is not None:
        return cli.upload(pptx, upload_cfg, args.overwrite)
    return 0


if __name__ == "__main__":
    sys.exit(_main())
