# TK Sport Kolovraty — Team Results Scraper

## What this does

Scrapes official match records from cesky-tenis.cz for any number of
team-per-season combinations, and reproduces the same statistics
you've been calculating by hand in Excel: Odehráno, Vyhráno,
Úspěšnost, Body (Suma + Váženo) per player, plus the Awards/Ocenění
ranking table, formatted with the same color/data-bar treatment as
your original sheet.

Which teams and seasons to track is controlled entirely by `teams.csv`
— no code editing needed for ordinary use.

## How to run it

1. Install Python 3 if you don't have it (any recent version is fine).
2. Install the two dependencies:
   ```
   pip install beautifulsoup4 openpyxl
   ```
3. Make sure `teams.csv` lists every team/season you want (see below).
4. From this folder, run:
   ```
   python3 main.py
   ```
5. It processes every row in `teams.csv` in turn, prints progress as
   it fetches, and writes one file per row to `/output`, named
   `{season_year}_{team_label}_extracted.xlsx`.
6. A summary at the end shows OK/FAILED for every row. One bad row
   (e.g. a stale ID) won't stop the others from completing.

## Adding a new season or team

Open `teams.csv` and add a row. Three things are needed, and only one
of them comes from the website:

- `url` — the team's page on the site, e.g.
  `https://cesky-tenis.cz/soutez/9029?druzstvoCislo=8`. `competition_id`
  and `team_number` are parsed straight out of this automatically, so
  there's nothing else to copy from the address bar.
- `season_year` — just a label used in the output filename, e.g. `2024`
- `team_label` — `A`, `B`, `C`, `dorost`, etc. — also just a filename label
- `competition_name` — free text, purely for your own readability when skimming the CSV

Find the team's standings page on the site the same way as before
(club page → team → that season), copy the URL out of the address
bar, paste it in. If it's easier, just paste the URL in chat and I'll
add the row directly.

The parser itself needs no changes for a new season *in the expected
case*, since the page structure is config-independent — but this is
an expectation, not a guarantee, which is exactly why testing a second
season is worth doing before trusting it long-term. Things that could
still surface for the first time: walkovers/scratches, a round with an
incomplete lineup, a season with a different number of rounds, or (less
likely, but possible) an older season rendered through a different
template.

## Files

- `main.py` — entry point, run this one. Loads `teams.csv`, loops over
  every row, writes one workbook per row.
- `teams.csv` — the list of team/season combinations to scrape. Edit
  this, not the Python, for ordinary use.
- `parser.py` — turns raw HTML into structured data (roster, rounds,
  per-position results).
- `aggregate.py` — turns structured data into per-player stats and the
  Awards rankings.
- `fetch.py` — polite HTTP fetching (delay + retries).
- `output/` — one `.xlsx` per `teams.csv` row, plus `dryrun_test.xlsx`,
  a preview built from a single verified round (no network needed),
  useful for checking formatting changes without a live run.

## Why it must run on YOUR computer, not in this chat

cesky-tenis.cz's `robots.txt` blocks automated tools, so Claude's own
sandbox cannot fetch live pages from it. This script runs on your
machine instead, where you have normal browser-level access to public
pages — the same pages you already open by hand. It's deliberately
slow (2 second delay between requests, roughly 9 requests per row in
`teams.csv`) to stay a polite, low-volume personal script even as more
rows get added.

## If something looks wrong

The most useful thing to send back is the console output (especially
any WARNING or FAILED lines) and which player/round disagrees with
your sheet, with both numbers.

## Next steps

1. Add a second season (e.g. 2024) as a new row and confirm it matches
   your existing sheet — this is what surfaces any season-specific
   edge cases mentioned above.
2. Add B tým and C tým as further rows once the season check is solid.
3. Add the cross-team combined view (the bottom block of your sheet
   that merges players who played for more than one team).
4. Only after that: scheduling, the dorost/žactvo categories, and the
   "interesting facts" detection you mentioned early on.
