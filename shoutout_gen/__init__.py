"""GM slide generators: shout-outs (spreadsheet -> non-overlapping, click-animated PPTX),
the question of the week, and the snipe game.

Modules (interface -> implementation, no cycles):
    google_sheet  optional: pull the form's responses tab from Google Sheets as CSV
    sheet    read the spreadsheet, pick the message column, order by timestamp
    metrics  measure/wrap text in the configured font (default Roboto) exactly as rendered
    layout   pack measured boxes onto slides with zero overlap
    deck     write the PPTX (textboxes + per-box click animation, the QOTW finale,
             the snipe slide's click reveal)
    qotw     tally the question-of-the-week column and draw its bar chart
    snipes   read snipebot's day-by-day workbook; the week's boards and the season race
    cli      the scripts' shared plumbing: sheet source, --since, Drive upload
    term     colour for the scripts' terminal output (plain when piped)
    drive    optional upload of the finished deck to Google Drive as Google Slides;
             owns the Google sign-in that google_sheet shares
"""
