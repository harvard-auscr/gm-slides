"""Chart the week's question-of-the-week answers as a slide.

    python qotw_chart.py --gm 4 --date 10.7 --since "9.30 20:00"
        -> pulls the form's responses sheet named in .env, counts the QOTW picks
           submitted since 9/30 20:00, writes output/GM #4 10.7---qotw.png and .pptx
    python qotw_chart.py input/form.csv --gm 4 --date 10.7 --since "9.30 20:00"
        -> the same from a downloaded export

The QOTW is the form column whose header starts "QOTW" (column D so far); its
answers are names, and the chart puts the most picked name on the left,
tapering to the right, under the question as the header words it. The PNG can
be pasted into any slide; the PPTX is that image on a club-theme slide and, if
``.env`` names a Google Drive folder, is uploaded there as Google Slides, the
same way the shout-out deck is.

The question changes every week but its old answers stay in the column, so pass
``--since`` with when the previous GM's slide was made, date and time, to count
only this week's.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from shoutout_gen import cli
from shoutout_gen.deck import TEMPLATE_PATH, write_image_slide
from shoutout_gen.metrics import FontNotFoundError, resolve_font
from shoutout_gen.qotw import output_name_for, read_tally, render_chart
from shoutout_gen.sheet import SheetFormatError


def _main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    cli.add_source_args(parser)
    parser.add_argument("--column", help="force the QOTW column, by header name or letter (default: the 'QOTW' header, else D)")
    parser.add_argument("--template", type=Path, default=TEMPLATE_PATH)
    parser.add_argument("--font", default=None, help="typeface for the chart (default Roboto)")
    parser.add_argument("--font-file", type=Path, default=None, help="TTF/OTF to draw the chart with (default: the bundled Roboto)")
    cli.add_upload_args(parser)
    args = parser.parse_args(argv)

    settings = cli.load_settings(args)
    if isinstance(settings, int):
        return settings
    upload_cfg, source = settings
    try:
        _, font_file = resolve_font(args.font, args.font_file)
    except FontNotFoundError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    sheet_path = cli.sheet_path(args, source)
    if isinstance(sheet_path, int):
        return sheet_path

    try:
        result, skipped = read_tally(sheet_path, args.column, args.since)
    except (SheetFormatError, FileNotFoundError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    if args.since is not None:
        print(f"since {args.since:%m/%d/%Y %H:%M}: {result.responses} answers kept, {skipped} older skipped")
    if not result.counts:
        print("error: no QOTW answers to chart" + (" in that window" if args.since is not None else ""), file=sys.stderr)
        return 2

    png = render_chart(result, args.out / output_name_for(args.gm, args.date, ".png"), font_file)
    pptx = write_image_slide(png, args.out / output_name_for(args.gm, args.date, ".pptx"), args.template)
    print(f"QOTW: {result.question}")
    for name, n in result.counts:
        print(f"  {n:>3}  {name}")
    print(f"{result.responses} responses -> {png}")
    print(f"slide -> {pptx}")

    if upload_cfg is not None:
        return cli.upload(pptx, upload_cfg, args.overwrite)
    return 0


if __name__ == "__main__":
    sys.exit(_main())
