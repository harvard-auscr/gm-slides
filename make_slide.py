"""Make the week's GM slide: the shout-outs, then the question of the week.

    python make_slide.py --gm 4 --date 10.7 --since "9.30 20:00"
        -> pulls the form's responses sheet named in .env and writes
           output/GM #4 10.7---shoutouts-qotw.pptx (+ the chart as ...---qotw.png)

One slide, all on clicks: every shout-out appears in the order it was
submitted (exactly as generate_shoutouts.py lays them out), then a full-slide
"QOTW: <question>" pane announces the question, then the chart of this week's
answers -- most picked name on the left -- replaces it. The same ``--since``
window applies to both. If nobody answered the QOTW in that window, the slide
is made without it and the summary says so.

If ``.env`` names a Google Drive folder the slide is uploaded there as Google
Slides, ready for File -> Import slides in the GM deck.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from shoutout_gen import cli, term
from shoutout_gen.deck import TEMPLATE_PATH, QotwFinale
from shoutout_gen.generate import AUTO_FONT_SIZES_PT, LAYOUT_MODES, generate
from shoutout_gen.metrics import FontNotFoundError, resolve_font
from shoutout_gen.qotw import read_tally, render_chart
from shoutout_gen.qotw import output_name_for as chart_name_for
from shoutout_gen.sheet import SheetFormatError, filter_since, read_shoutouts, split_placeholders


def output_name_for(gm_number: int, date: str) -> str:
    """"GM #4 10.7---shoutouts-qotw.pptx", beside the shout-out-only deck's name."""
    return f"GM #{gm_number} {date}---shoutouts-qotw.pptx"


def _font_size(value: str) -> float | None:
    """argparse type: "auto" -> None (size ladder), otherwise a point size."""
    if value.lower() == "auto":
        return None
    size = float(value)
    if size <= 0:
        raise argparse.ArgumentTypeError("font size must be positive")
    return size


def _main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    cli.add_source_args(parser)
    parser.add_argument(
        "--font-size", type=_font_size, default=None,
        help=f"shout-out point size, or 'auto' (default) to try {', '.join(f'{s:g}' for s in AUTO_FONT_SIZES_PT)} until one slide fits",
    )
    parser.add_argument("--layout", choices=LAYOUT_MODES, default="auto")
    parser.add_argument("--seed", type=int, default=0, help="change for a different arrangement")
    parser.add_argument("--text-column", help="force the shout-out column (header name)")
    parser.add_argument("--time-column", help="force the timestamp column (header name)")
    parser.add_argument("--qotw-column", help="force the QOTW column, by header name or letter (default: the 'QOTW' header, else D)")
    parser.add_argument("--no-qotw", action="store_true", help="shout-outs only, like generate_shoutouts.py")
    parser.add_argument("--template", type=Path, default=TEMPLATE_PATH)
    parser.add_argument("--font", default=None, help="typeface for the slide (default Roboto)")
    parser.add_argument("--font-file", type=Path, default=None, help="TTF/OTF matching --font, used to measure and draw (default: the bundled Roboto)")
    cli.add_upload_args(parser)
    args = parser.parse_args(argv)

    settings = cli.load_settings(args)
    if isinstance(settings, int):
        return settings
    upload_cfg, source = settings
    try:
        font_name, font_file = resolve_font(args.font, args.font_file)
    except FontNotFoundError as exc:
        term.error(str(exc))
        return 2
    sheet_path = cli.sheet_path(args, source)
    if isinstance(sheet_path, int):
        return sheet_path

    # --- shout-outs: the same reading and filtering as generate_shoutouts.py ---
    try:
        shoutouts = read_shoutouts(sheet_path, args.text_column, args.time_column)
    except (SheetFormatError, FileNotFoundError) as exc:
        term.error(str(exc))
        return 2
    shoutouts, placeholders = split_placeholders(shoutouts)
    if placeholders:
        term.detail(f"skipped {len(placeholders)} placeholder row(s) with no real message")
    if args.since is not None:
        shoutouts, skipped = filter_since(shoutouts, args.since)
        term.detail(f"since {args.since:%m/%d/%Y %H:%M}: {len(shoutouts)} shout-outs kept, {skipped} older skipped")
    if not shoutouts:
        term.error("no shout-outs to put on the slide" + (" in that window" if args.since is not None else ""))
        return 2

    # --- question of the week ---
    finale = None
    if not args.no_qotw:
        try:
            result, skipped = read_tally(sheet_path, args.qotw_column, args.since)
        except SheetFormatError as exc:
            term.error(f"{exc} (or pass --no-qotw)")
            return 2
        print()
        if result.counts:
            png = render_chart(result, args.out / chart_name_for(args.gm, args.date, ".png"), font_file)
            finale = QotwFinale(result.question, png)
            cli.print_tally(result)
            term.made(f"{result.responses} QOTW responses", png)
        else:
            term.note("no QOTW answers in this window; the slide has the shout-outs only")

    out = generate(
        [s.text for s in shoutouts],
        args.out / output_name_for(args.gm, args.date),
        font_size_pt=args.font_size,
        layout_mode=args.layout,
        seed=args.seed,
        template_path=args.template,
        font_name=font_name,
        font_file=font_file,
        qotw=finale,
    )
    stamped = [s.timestamp for s in shoutouts if s.timestamp is not None]
    clicks = len(shoutouts) + (2 if finale else 0)
    print()
    term.made(f"{len(shoutouts)} shout-outs{' + QOTW' if finale else ''}", out.output, f" ({clicks} clicks)")
    if stamped:
        term.detail(f"shout-outs submitted {min(stamped):%m/%d/%Y %H:%M} -> {max(stamped):%m/%d/%Y %H:%M}")
    term.detail(f"slides: {out.slide_count} (shout-outs per slide: {list(out.per_slide_counts)}), font {out.font_size_pt:g}pt")
    if out.slide_count > 1:
        term.note("the shout-outs did not fit on one slide; the QOTW follows them on the last slide")

    if upload_cfg is not None:
        return cli.upload(out.output, upload_cfg, args.overwrite)
    return 0


if __name__ == "__main__":
    sys.exit(_main())
