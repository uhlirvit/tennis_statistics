# TK Sport Kolovraty — Team Results Scraper

## What this does

Scrapes official match records from cesky-tenis.cz for any number of
team-per-season combinations, and reproduces the same statistics
you've been calculating by hand in Excel: Odehráno, Vyhráno,
Úspěšnost, Body (Suma + Váženo) per player, plus the Awards/Ocenění
ranking table, formatted with the same color/data-bar treatment as
your original sheet.

Which teams and seasons to track is controlled entirely by `teams.csv`
— no code editing needed for ordinary use. Rows that share the same
season and category are combined into one workbook, with one set of
four sheets per team (mirroring how your original Excel kept every
team for a season in a single tab).

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
5. Rows are grouped by `(season_year, category)`. Each group becomes
   one file in `/output`, named `{season_year}_{category}.xlsx`, with
   one four-sheet group per team inside it: "{team_label} - Player
   Summary", "{team_label} - Match Grid", "{team_label} - Match Log",
   "{team_label} - Awards".
6. A summary at the end shows OK/FAILED for every row, nested under
   its season/category. One bad row won't stop the others in the same
   group from being written.

## Adding a new season or team

Open `teams.csv` and add a row. Five columns, and only one of them
comes from the website:

- `url` — the team's page on the site, e.g.
  `https://cesky-tenis.cz/soutez/9029?druzstvoCislo=8`. `competition_id`
  and `team_number` are parsed straight out of this automatically, so
  there's nothing else to copy from the address bar.
- `season_year` — just a label, e.g. `2024`.
- `category` — `dospeli` / `dorost` / `zactvo` / etc. Defaults to
  `dospeli` if left blank. Rows only combine into one file when both
  `season_year` AND `category` match — this is what keeps a future
  youth category from landing in the same file as the adult teams.
- `team_label` — `A` / `B` / `C` / etc. Becomes the sheet-name prefix.
- `competition_name` — free text, purely for your own readability when
  skimming the CSV.

Find the team's standings page on the site the same way as before
(club page → team → that season), copy the URL out of the address
bar, paste it in. If it's easier, just paste the URL in chat and it
can be added as a row directly.

The parser itself needs no changes for a new season or team *in the
expected case*, since the page structure is config-independent — but
this is an expectation, not a guarantee, which is why testing each new
row is worth doing before trusting it long-term. One real example
already found this way: lower teams can be listed under a suffixed
name (e.g. "TK Sport Kolovraty C") rather than the bare club name —
handled automatically now, but it only surfaced by actually running a
new row. Other things that could still surface for the first time:
walkovers/scratches, a round with an incomplete lineup, or a season
with a different number of rounds.

## Files

- `main.py` — entry point, run this one. Loads `teams.csv`, groups
  rows by season+category, writes one workbook per group.
- `teams.csv` — the list of team/season combinations to scrape. Edit
  this, not the Python, for ordinary use.
- `parser.py` — turns raw HTML into structured data (roster, rounds,
  per-position results, and the team's own display name).
- `aggregate.py` — turns structured data into per-player stats and the
  Awards rankings.
- `fetch.py` — polite HTTP fetching (delay + retries).
- `output/` — one `.xlsx` per `(season_year, category)` group, plus
  `dryrun_test.xlsx`, a preview built from a single verified round (no
  network needed), useful for checking formatting changes without a
  live run.

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

Season and team coverage is now validated: 2025 A, 2024 A, and 2025 C
all run cleanly, including a real bug found and fixed along the way
(lower teams using a suffixed display name).

1. Add B tým, and any other remaining teams/categories, as further
   `teams.csv` rows.
2. Add the all-club sheet (the bottom block of your original sheet
   that merges players who played for more than one team in a season)
   — this slots in as an additional sheet-group in the same per-season
   workbook, since all teams' data for a season is already gathered
   together by the time the file gets written.
3. Only after that: the multi-season history table, which will need
   its own mechanism since it spans beyond any single season's file —
   a separate design conversation when we get there.
4. Further out: scheduling, and the "interesting facts" detection
   mentioned early on.
