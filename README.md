# HAUSCR shout-out slides

Turns the weekly shout-out form export into a PowerPoint slide: every shout-out
is its own textbox, **nothing overlaps**, and each one **appears on its own
click** in the order it was submitted — the same look as the club's past GM
slides, minus the collisions.

```
python make_slide.py --gm 4 --date 10.7 --since "9.30 20:00"
# -> pulls the form's responses sheet (link in .env), keeps what came in since last GM,
#    writes output/GM #4 10.7---shoutouts-qotw.pptx: the shout-outs, then the
#    question of the week and a chart of its answers, all on clicks
python generate_shoutouts.py --gm 26 --date 9.24 --since "9.17 20:00"
# -> shout-outs only: output/GM #26 9.24---shoutouts.pptx
python generate_shoutouts.py input/shoutouts.csv --gm 25 --date 9.10
# -> the same from a downloaded export
```

## Setup

Python 3.10+.

```
git clone <this repo>
cd hauscr-shoutouts
python -m venv .venv && .venv\Scripts\activate      # macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
python generate_shoutouts.py samples/example_form.csv --gm 1 --date 9.17   # smoke test
```

## Weekly use

With the responses sheet connected (one-time, see [Google Sheet
source](#google-sheet-source-optional)):

1. Run `python make_slide.py --gm <n> --date <m.d> --since "<previous GM's m.d> <HH:MM>"`.
   The tab behind the form is pulled into `input/GM #<n> <m.d>---form.csv`,
   responses submitted before `--since` are left out (the form collects all
   year), and the slide is built from the rest. Give `--since` the time the
   previous slide was made, not just the day: shout-outs sent on GM day before
   the meeting were already on last week's slide.
2. Open `output/GM #<n> <m.d>---shoutouts-qotw.pptx`, check it, paste the slide into
   the GM deck — or let the optional [Drive upload](#google-drive-upload-optional)
   put a Google Slides copy in a shared folder for **File → Import slides**.

`generate_shoutouts.py` takes the same options and makes the shout-outs alone.

In a terminal the summary is colour-coded: the files made in green with their
paths and links in cyan, notes in yellow, errors in red, and the bookkeeping
(rows skipped, the `--since` window) in grey. Piped or redirected output stays
plain text; set `NO_COLOR=1` to turn colour off, `FORCE_COLOR=1` to force it on.

### Question of the week

The form's QOTW column (header starting `QOTW`, e.g. `QOTW [9/30]: Which board
member is the greenest flag?`; column D so far) is answered with names.
`make_slide.py` adds two clicks after the last shout-out: a full-slide
**QOTW: <question>** pane (covering the shout-outs), then a bar chart of this
week's answers — most picked name on the left, tapering right, count over each
bar. The question is taken from the header, minus its `QOTW [date]:` tag. A
checkbox question's multiple picks are each counted; spellings that differ only
in case or spacing are merged. With no answers in the window the slide is made
without the QOTW, and the summary says so; `--no-qotw` leaves it off on purpose,
and `--qotw-column "<header>"` (or a letter, `--qotw-column E`) overrides the
column.

The chart on its own, as a PNG to paste anywhere plus a one-slide `.pptx`:

```
python qotw_chart.py --gm 4 --date 10.7 --since "9.30 20:00"
# -> output/GM #4 10.7---qotw.png and output/GM #4 10.7---qotw.pptx
```

Without the connection: export the responses yourself (**File → Download →
CSV** or `.xlsx`), save the file in `input/` (git-ignored) and pass it in place
of the link: `python generate_shoutouts.py input/<file> --gm <n> --date <m.d>`.
A sheet link on the command line works too.

The sheet needs a timestamp column and a message column; they are detected from
the headers (`Timestamp`, `Shout-out`, …). If detection guesses wrong:
`--text-column "<header>"` / `--time-column "<header>"`.
Rows whose message is only punctuation (a `/` typed to get past the required
box and answer the form's other questions) are skipped, and the summary says
how many.

| option | default | meaning |
|---|---|---|
| `--font-size` | `auto` | 12pt, stepping down to 11 then 10 until everything fits one slide; a number pins it (a huge week then spills to a second slide and the summary says so) |
| `--layout` | `auto` | `compact` (rows with jitter) if it fits one slide, else `dense` |
| `--seed` | `0` | different number = different arrangement |
| `--out` | `output` | output folder |
| `--since` | all rows | keep only shout-outs submitted at or after this moment: `"9.17 21:00"`, `9.17` (midnight) or `2026-09-17 21:00`; a bare `m.d` means the most recent such day, and is refused if that is more than six months back |
| `--input-dir` | `input` | where a tab pulled from Google Sheets is saved as CSV before it is read |
| `--no-upload` | off | skip the Google Drive upload this run even if `.env` enables it (a connected sheet is still read) |
| `--overwrite` | off | replace a same-name Google Slides file in the Drive folder without asking (default: ask; no answer keeps it) |
| `--env-file` | `.env` in the repo | where the upload settings live |
| `--template` | `template/shoutouts_template.pptx` | the club-theme template every deck is built from |
| `--font` | `Roboto` | typeface for the slide text, looked up by name in `fonts/` and the Windows font folders |
| `--font-file` | the bundled Roboto | TTF/OTF used to measure and wrap the text; with only this flag, the typeface name is read from the file |

A custom font must also exist wherever the deck is presented — Google Slides
only ships Google Fonts, and PowerPoint uses whatever is installed on the
machine. Text is wrapped using the font file, so `--font` and `--font-file`
must describe the same font or the renderer will wrap differently than
predicted.

Optional, Windows + PowerPoint only — open the deck in real PowerPoint, check the
*rendered* text bounds pairwise, count the animations, export PNGs:

```
python verify_render.py "output/GM #25 9.10---shoutouts.pptx" --png-dir renders
```

On a `make_slide.py` slide the QOTW pane and chart are left out of the overlap
check (they cover the slide on purpose) and must be the last two clicks.

## Google Sheet source (optional)

Skip the weekly download: put the responses sheet's link in `.env` as
`SHOUTOUT_SHEET` (or pass it on the command line in place of the file) and the
generator fetches the tab itself. The tab in the link (`gid=…`) is the one that
is read, through the sheet's own CSV export — byte for byte what **File →
Download → CSV** gives you — and the download is kept as
`input/GM #<n> <m.d>---form.csv`, so a run can always be repeated offline from
that file.

It signs in the same way as the Drive upload (steps 1–2 [below](#google-drive-upload-optional):
the OAuth client in `.env`, the cached token in `.google/`; `DRIVE_FOLDER` may
stay empty) and needs nothing else in the Cloud console — not the Sheets API.
The signed-in Google account must be able to open the sheet. If the download
fails, nothing is generated and the error says why (wrong tab, no access,
expired sign-in); download the tab by hand and pass the file to carry on.

The form collects responses all year, so pass `--since` with when the previous
GM's slide was made, date and time (`"9.17 20:00"`), to keep only the new ones.
A bare date means midnight, so shout-outs sent earlier that day, already on last
week's slide, would come back. A bare date more than six months back is refused
as a likely typo; give the full date (`2025-10-29`) if you mean it. The summary
prints how many rows were left out and the submission span that made the slide.

## Google Drive upload (optional)

The GM decks live in Google Slides, so the generator can drop each new deck
straight into a Drive folder, already converted to a Google Slides file — after
that it's **File → Import slides** in the GM deck. Nothing is uploaded unless
you opt in.

Before you do: the consent screen asks for **full Google Drive access** (the
`drive` scope), because the narrower `drive.file` scope cannot see a folder you
picked by pasting its link. The cached token in `.google/` carries that access,
so keep it on the machine that runs the generator and delete it when you stop
using the upload.

1. In the [Google Cloud console](https://console.cloud.google.com/) create a
   project, enable the **Google Drive API**, set up the OAuth consent screen
   (External; add yourself as a test user) and create an **OAuth client ID** of
   type *Desktop app*.
2. Copy `.env.example` to `.env` (git-ignored) and fill in `DRIVE_FOLDER` (paste
   the folder link) plus either `DRIVE_CLIENT_ID` + `DRIVE_CLIENT_SECRET` or
   `DRIVE_CLIENT_FILE` pointing at the downloaded client JSON.
3. Run the generator as usual. The first upload opens a browser for consent
   once; the sign-in is cached in `.google/` (git-ignored) after that.

`DRIVE_UPLOAD=off` in `.env` (or `--no-upload`) skips the upload without
touching the rest of the config; the same names as real environment variables
override the file. If an upload fails the deck is still in `output/` — retry
with `python upload_to_slides.py "output/GM #25 9.10---shoutouts.pptx"`.

Regenerating a week: if a Google Slides file with the deck's name is already in
the folder, the generator asks on the terminal before **overwriting it in
place** — same file, same link, so a slide already imported into the GM deck
is unaffected and a shared link keeps working. Answering anything but `y`
(or running without a terminal) keeps the existing file and leaves the new
deck in `output/`; `--overwrite` answers yes without asking. A second copy is
never created.

## Data: what goes where

| folder | contents | in git? |
|---|---|---|
| `input/` | weekly form exports (`.csv` / `.xlsx`), including the CSV pulled from a connected sheet (`GM #<n> <m.d>---form.csv`) | no |
| `dataset/` | historical GM decks, named `General Meeting #<n> <m.d>.pptx` — only needed to rebuild the corpus/template or run the calibration test | no |
| `output/` | generated decks; `output/scraped/` and `output/shoutouts_corpus.*` when the corpus tools are run | no |
| `samples/` | `example_form.csv` (synthetic, safe to publish); CSVs built from real shout-outs stay local | example only |
| `template/` | `shoutouts_template.pptx` — the club theme with a single `ONE_COLUMN_TEXT` layout | yes |
| `fonts/` | `Roboto-Regular.ttf` (Apache 2.0, see `fonts/LICENSE.txt`) used to measure text | yes |
| `.env`, `.google/` | sheet link, Drive upload settings and the cached Google sign-in (template: `.env.example`) | example only |

Everything derived from members' submissions is ignored by default because it
contains names, photos and in-jokes. If this repo is private and you want the
history versioned, delete the corresponding lines from `.gitignore`.

## Corpus tools (one-off, need `dataset/`)

```
python scrape_shoutouts.py     # dataset/ -> output/scraped/GM #<n> <m.d>---shoutouts.pptx
python build_corpus.py         # output/scraped/ -> output/shoutouts_corpus.{xlsx,csv}
python make_template.py        # output/scraped/<deck> -> template/shoutouts_template.pptx
```

## How it works

| step | module | what it does |
|---|---|---|
| pull | `shoutout_gen/google_sheet.py` | optional: download the responses tab of the connected Google Sheet as CSV into `input/`, with the upload's sign-in |
| read | `shoutout_gen/sheet.py` | load CSV/XLSX, detect columns, order by timestamp, apply `--since` |
| measure | `shoutout_gen/metrics.py` | wrap each message in Roboto exactly as the slide will, at its own seeded wrap width (2.4-4.2 in, so wrapped boxes do not line up into columns), with a 10% width slack calibrated on 531 real boxes so the renderer never wraps earlier than predicted |
| pack | `shoutout_gen/layout.py` | place boxes on a 0.05in raster using prefix-sum feasibility, then stretch the layout to span the slide (scaling positions up can never create an overlap) |
| write | `shoutout_gen/deck.py` | textboxes + one click-Appear animation per box, from the template |
| upload | `shoutout_gen/drive.py` | optional: push the deck to a Drive folder as Google Slides (off unless `.env` says so) |

Why we trust "no overlaps": boxes are placed only on cells proven free (and
re-checked pairwise before saving); the text model never under-predicts line
count on the historical corpus; and `verify_render.py` confirms the rendered
text bounds in real PowerPoint.

## Tests

```
python -m pytest tests -q
```

Layout, measurement, sheet and deck tests always run. The corpus calibration and
end-to-end tests need `dataset/` + the corpus tools' output and skip otherwise.
CI (`.github/workflows/tests.yml`) runs the suite and the quickstart on 3.11/3.12.

## License

MIT (see `LICENSE`). Roboto is bundled under the Apache License 2.0
(`fonts/LICENSE.txt`).
