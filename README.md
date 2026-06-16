# TK Sport Kolovraty — Team Results Scraper (Step 1: A tým, 2025)

## What this does

Scrapes the official match records for one team's one season from
cesky-tenis.cz and reproduces the same statistics you've been
calculating by hand in Excel: Odehráno, Vyhráno, Úspěšnost, Body
(Suma + Váženo) per player.

## Why it must run on YOUR computer, not in this chat

cesky-tenis.cz's `robots.txt` blocks automated tools, so Claude's own
sandbox cannot fetch live pages from it. This script is built to run
on your machine instead, where you have normal browser-level access to
public pages — the same pages you already open by hand. It is
deliberately slow (2 second delay between each of the 8 requests this
needs) to be a polite, low-volume personal script.

## What has already been verified (without touching the live site)

Using the two pages you pasted earlier (the team page and one match
record), I tested every parsing and aggregation function against real
markup. The result matched your existing `2025_dospeli` sheet exactly
for round 1 (vs I.ČLTK): Marian 0/0, Milan 1/1, Víťa 1/1, Béďa 0/0,
Kája 1/1, Jessica 0/1 — and the round order/opponents (ČLTK → HAMR →
Topolka → Hradecká → Vysočany → Modřany → Aritma) match your
"Mistrák" row exactly. `output/dryrun_test.xlsx` shows what the real
output will look like, built from that one verified round.

What's *not* yet tested: actually fetching all 7 rounds live, since
that requires network access this sandbox doesn't have.

## How to run it

1. Install Python 3 if you don't have it (any recent version is fine).
2. Install the two dependencies:
   ```
   pip install beautifulsoup4 openpyxl
   ```
3. From this folder, run:
   ```
   python3 main.py
   ```
4. It will print progress for each of the 7 rounds as it fetches them,
   then write `output/2025_A_tym_extracted.xlsx`.
5. Open that file and compare the "Player Summary" sheet against your
   `2025_dospeli` A tým block. They should match exactly.

## If something looks wrong

The most useful thing to send back is:
- the console output (especially any "WARNING" lines)
- which player/round disagrees with your sheet, and what the two
  numbers are

Likely sources of mismatch we haven't seen yet in the data: walkovers
("kontumačně"/scratch matches), a round where Kolovraty didn't field a
full team, or a player who only appears mid-season. The parser is
written to not crash on these, but I can't promise it handles them
*correctly* until we see one and check.

## Files

- `main.py` — entry point, run this one. Has a small CONFIG block at
  the top (competition ID, team number, club name) in case you want to
  point it at a different team/season later.
- `parser.py` — turns raw HTML into structured data (roster, rounds,
  per-position results).
- `aggregate.py` — turns structured data into the per-player stats.
- `fetch.py` — polite HTTP fetching (delay + retries).
- `output/dryrun_test.xlsx` — preview built from the one verified
  round, so you can see the format before running anything live.

## Next steps once this is verified

1. Confirm all 7 rounds match your sheet exactly.
2. Find the `druzstvoCislo` for B tým and C tým (visible in the
   standings table links on their competition pages) and point the
   same script at them.
3. Add the cross-team combined view (the bottom block of your sheet).
4. Only then: scheduling, the dorost/žactvo categories, and the
   "interesting facts" detection you mentioned.
