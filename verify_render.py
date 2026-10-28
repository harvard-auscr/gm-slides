"""Open a generated deck in real PowerPoint and check what it actually renders.

Why: python-pptx can only prove the *boxes* we wrote don't overlap. Whether the
*text* overlaps depends on how the renderer wraps it, so this asks PowerPoint
(via COM) for every text run's rendered bounds and checks those pairwise, plus
that each shape's animation exists and that nothing leaves the slide. It also
exports a PNG per slide for eyeballing.

Caveat: PowerPoint renders Roboto here through Office's cloud fonts and the
template's embedded copy; on a machine with neither it substitutes another
font, and a pass is then a pass under a *different* font than the layout was
measured for -- still a useful robustness check, but not the exact target.
Google Slides (where the deck is presented) has Roboto natively.

Usage:
    python verify_render.py "output/GM #25 9.10---shoutouts.pptx" [--png-dir folder]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from shoutout_gen.deck import QOTW_CHART_NAME, QOTW_PANE_NAME
from shoutout_gen.layout import overlapping_pairs


def verify(pptx_path: Path, png_dir: Path | None = None) -> list[str]:
    """Return a list of problems (empty == pass). Prints a per-slide summary."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    problems: list[str] = []
    pres = app.Presentations.Open(str(pptx_path.resolve()), ReadOnly=True, Untitled=False, WithWindow=False)
    try:
        page_w, page_h = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        for s in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(s)
            text_rects, box_rects = [], []
            finale = [shape.Name for shape in slide.Shapes if shape.Name in (QOTW_PANE_NAME, QOTW_CHART_NAME)]
            for shape in slide.Shapes:
                # The QOTW pane and chart cover the whole slide on purpose; only shout-outs must not overlap.
                if not shape.HasTextFrame or shape.Name in finale:
                    continue
                tr = shape.TextFrame.TextRange
                text_rects.append((tr.BoundLeft, tr.BoundTop, tr.BoundWidth, tr.BoundHeight))
                box_rects.append((shape.Left, shape.Top, shape.Width, shape.Height))
                if shape.Left < 0 or shape.Top < 0 or shape.Left + shape.Width > page_w or shape.Top + shape.Height > page_h:
                    problems.append(f"slide {s}: shape {shape.Name!r} leaves the page")
                # If the renderer wrapped a line we didn't expect, the text gets
                # taller than the box we packed -- that is the failure that
                # produces overlaps. (BoundWidth is not compared: PowerPoint
                # pads it by a constant ~3pt of glyph overhang, and an over-wide
                # line would show up as extra height anyway.)
                if tr.BoundHeight > shape.Height + 0.5:
                    problems.append(
                        f"slide {s}: text of {tr.Text[:30]!r} renders {tr.BoundHeight:.1f}pt tall, "
                        f"taller than its {shape.Height:.1f}pt box (unexpected wrap)"
                    )
            text_hits = overlapping_pairs(text_rects)
            box_hits = overlapping_pairs(box_rects)
            for i, j in text_hits:
                problems.append(f"slide {s}: rendered text overlaps between shapes {i} and {j}")
            sequence = slide.TimeLine.MainSequence
            anims = sequence.Count
            if anims != len(text_rects) + len(finale):
                problems.append(f"slide {s}: {anims} animations for {len(text_rects)} shout-outs + {len(finale)} QOTW shapes")
            elif finale and [sequence(i).Shape.Name for i in range(anims - 1, anims + 1)] != [QOTW_PANE_NAME, QOTW_CHART_NAME]:
                problems.append(f"slide {s}: the QOTW pane and chart are not the last two clicks, in that order")
            print(
                f"slide {s}: {len(text_rects)} shout-outs{' + QOTW pane, chart' if finale else ''}, {anims} click animations, "
                f"text overlaps={len(text_hits)}, box overlaps={len(box_hits)}"
            )
            if png_dir is not None:
                # PowerPoint resolves relative paths against ITS working directory,
                # not ours, so a relative --png-dir must be made absolute first.
                png_dir = png_dir.resolve()
                png_dir.mkdir(parents=True, exist_ok=True)
                slide.Export(str(png_dir / f"{pptx_path.stem}_s{s}.png"), "PNG", 1920, 1080)
    finally:
        pres.Close()
        app.Quit()
    return problems


def _main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--png-dir", type=Path, default=None)
    args = parser.parse_args(argv)
    problems = verify(args.pptx, args.png_dir)
    for p in problems:
        print("PROBLEM:", p)
    print("PASS" if not problems else f"FAIL ({len(problems)} problems)")
    return 0 if not problems else 1


if __name__ == "__main__":
    sys.exit(_main())
