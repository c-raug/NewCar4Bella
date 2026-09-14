# Self-refreshing Google Sheet

`Code.gs` turns the spreadsheet into a live tool: one menu click re-pulls CARFAX,
rescores everything, and rewrites the tabs — keeping your criteria, your weights
and every note you have typed.

## One-time setup

1. Get the data in first. In **New Car for Bella**, go to
   **File → Import → Upload**, choose `output/Bella_Car_Search.xlsx`, and pick
   **Replace spreadsheet**. The URL stays the same.
2. **Extensions → Apps Script**. Delete the stub `myFunction` and paste all of
   `Code.gs`. Click **Save**.
3. Back on the sheet, reload the browser tab. A **Car Search** menu appears to
   the right of Help.
4. **Car Search → Refresh from CARFAX now**. Google asks for authorization the
   first time — it needs permission to edit this spreadsheet and to fetch
   carfax.com. Review and allow.
5. Optional: **Car Search → Turn on nightly auto-refresh** runs it daily around 5am.

## What a refresh preserves

| Kept | How |
|---|---|
| Search Criteria tab | Read before the rebuild and never overwritten. The tab is only created if missing. |
| Scoring Weights tab | Same. |
| Status / Notes / Contacted? | Read from both All Cars and Shortlist, merged by VIN, written back onto the same car. A hidden `_NotesStore` tab is the durable record. |
| Notes on a car that sold | The car is kept at the bottom of All Cars marked `DELISTED` rather than disappearing. |

The three note columns are shaded yellow on both tabs. Edit them wherever you
like — the merge reads both, and a non-empty value wins.

## What a refresh rebuilds

All Cars, Shortlist, Top 10 Compare, Model Benchmarks, and the Start Here counts.
Score, Rank and "Meets All" are written as live formulas pointing at the Criteria
and Weights tabs, so changing a yellow cell re-ranks the sheet instantly without
a refresh.

## Changing the search

Edit `CONFIG` at the top of `Code.gs` — ZIP, radius, the five target models,
mileage caps, default weights. Price cap and minimum year come from the Search
Criteria tab at run time, so change those on the sheet, not in code.

## Known risk

The script calls CARFAX's search endpoint from Google's servers with
`UrlFetchApp`. That call works from a normal machine, but it has not been
possible to test it from Google's IP ranges from here — CARFAX may rate-limit or
block it. If that happens you get an alert naming the HTTP status and **the sheet
is left untouched**; the refresh aborts before writing anything. Fall back to
re-running the Python pipeline in `scripts/` and re-importing the `.xlsx`.

## Scoring

Identical to the Python pipeline in `scripts/` — the port was verified
row-for-row against it across all 262 listings and all six sub-scores, with no
differences.
