# Paumenchon

A small web spreadsheet that mimics the basics of Google Sheets, built as the first
of several planning apps (planner, trips, Christmas gifts, to-do lists) selectable
from the top bar.

The default workbook mirrors the structure of the shared family calendar Google
Sheet: one sheet per year, months laid out in blocks of 9 columns side by side
(2 per band), each block = [day][weekday][one column per family member]
[Événements][Info supp][spacer], with day rows 1–31 and weekday letters
pre-filled for the correct year. The member columns are labeled `Membre 1`–`Membre 4`
(placeholders, since this repo is public) — rename them to your names by editing
the header cells. Regenerate the seed with `python3 tools/make_seed.py`.

## Run it

No build step, no dependencies. Either:

- open `index.html` directly in a browser, or
- serve the folder: `python3 -m http.server 8000` then open http://localhost:8000

## Features

- Editable grid: double-click or Enter to edit, arrows/Tab to move, Esc to cancel
- Multiple sheets (tabs at the bottom), add / rename / delete sheets
- Insert / delete rows and columns
- Bold, text color, cell fill (applied to the selection; Shift-click or click a
  row/column header to select a range)
- Formula bar; formulas like `=A1+B2`, `=SUM(A1:B5)`, `=AVG()`, `=MIN()`, `=MAX()`
- Autosaves to browser localStorage
- Import/export: CSV (replaces the active sheet) and JSON (whole workbook)

## Moving your real data in

Google Sheets can export one tab at a time as CSV (File → Download → CSV, or
per-tab via the `export?format=csv&gid=<tabId>` URL). Use **Import CSV/JSON** in
the toolbar: drop the CSV on the right sheet tab. Cell colors are not carried by
CSV; re-apply them with the Fill button. Alternatively paste the CSV of each tab
here and the workbook JSON can be generated for a one-shot **Import JSON**.

## Data model

A workbook is plain JSON: `{ sheets: [{ name, rows, cols, cells: { "r,c": { v, b, color, bg } } }] }`.
That makes it trivial to sync later (see roadmap).

## Roadmap

1. This static app with import/export (current state) — each person exports/imports JSON.
2. Two-user version: tiny backend (or Firebase/Supabase) storing the workbook JSON,
   auto-save on edit, refresh on the other side — enough for two people.
3. Live collaboration (WebSockets + per-cell last-write-wins or CRDT), then the
   remaining apps (trips, gifts, to-dos) on the same hub.
