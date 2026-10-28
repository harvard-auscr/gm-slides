"""Shout-out slide generator: spreadsheet -> non-overlapping, click-animated PPTX.

Modules (interface -> implementation, no cycles):
    google_sheet  optional: pull the form's responses tab from Google Sheets as CSV
    sheet    read the spreadsheet, pick the message column, order by timestamp
    metrics  measure/wrap text in the configured font (default Roboto) exactly as rendered
    layout   pack measured boxes onto slides with zero overlap
    deck     write the PPTX (textboxes + per-box click animation, the QOTW finale)
    qotw     tally the question-of-the-week column and draw its bar chart
    cli      the scripts' shared plumbing: sheet source, --since, Drive upload
    drive    optional upload of the finished deck to Google Drive as Google Slides;
             owns the Google sign-in that google_sheet shares
"""
