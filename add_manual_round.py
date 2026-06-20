"""
add_manual_round.py — Add a not-yet-published round by hand, then
re-run everything downstream exactly as main.py would.

Use case: the season's last match is finished and you know the result,
but cesky-tenis.cz hasn't published it yet. Fill in manual_round.csv
with the match's rows (one row per position per OUR player — same shape
as the Match Log sheet), then run this script. It will:

  1. Re-scrape every team in that season/category from teams.csv (so it
     has all the rounds the website DOES already have)
  2. Merge in your manual round on top
  3. Recompute Player Summary, Match Grid, Match Log, Awards — for every
     team in the group AND the All Club sheet
  4. Re-write the season's .xlsx
  5. If category is "dospeli": update history_store.json and
     club_history_dospeli.xlsx too, same as a normal main.py run

It does NOT regenerate the chart/HTML scripts (visualize.py,
season_awards.py, history_charts.py) — run those yourself afterward
once you're happy with the xlsx. Keeping this script single-purpose
means you can re-run it safely without re-triggering chart generation
every time.

Setup:
    python3 add_manual_round.py --init
        → creates manual_round.csv with header + example rows + comments

Usage:
    # edit manual_round.csv with your match's real results, then:
    python3 add_manual_round.py

    python3 add_manual_round.py --file my_round.csv
"""

import argparse
import csv
import sys
from pathlib import Path

from main import (
    load_teams, group_teams, process_one_team,
    write_season_workbook, update_history_store, write_history_workbook,
    TEAMS_CONFIG_PATH, OUTPUT_DIR, DEFAULT_CATEGORY, HISTORY_STORE_PATH,
)
from aggregate import compute_player_summary, order_by_roster, build_wide_grid

MANUAL_CSV_PATH = "manual_round.csv"
MANUAL_CSV_HEADER = [
    "season_year", "category", "team_label",
    "kolo", "datum", "souper", "pozice", "typ", "jmeno", "partner", "vysledek",
]

TEMPLATE = """\
# manual_round.csv — fill in the last match's results here, then run:
#     python3 add_manual_round.py
#
# Same columns as the Match Log sheet in the xlsx, so you can read them
# straight off a scoresheet:
#   season_year  e.g. 2026
#   category     dospeli / dorost / mladsi / baby / ... (blank = dospeli)
#   team_label   A / B / C / D / blank (blank = the category's one team)
#   kolo         round number, e.g. 8
#   datum        e.g. 28.06.2026
#   souper       opponent club name, e.g. TJ Sokol Foo
#   pozice       position number, e.g. 1
#   typ          singl / debl
#   jmeno        player name — MUST match exactly how they're written
#                everywhere else (incl. diacritics), or they'll show up
#                as a separate "new" player instead of being merged in
#   partner      doubles partner's name (blank for singl)
#   vysledek     W / L
#
# One row per OUR PLAYER per position. A doubles position = 2 rows
# (one per partner, each listing the other as "partner").
#
# Example (delete before filling in your own):
# season_year,category,team_label,kolo,datum,souper,pozice,typ,jmeno,partner,vysledek
# 2026,dospeli,A,8,28.06.2026,TJ Sokol Foo,1,singl,Roušar Petr,,W
# 2026,dospeli,A,8,28.06.2026,TJ Sokol Foo,2,singl,Uhlíř Vít,,L
# 2026,dospeli,A,8,28.06.2026,TJ Sokol Foo,3,debl,Roušar Petr,Uhlíř Vít,W
# 2026,dospeli,A,8,28.06.2026,TJ Sokol Foo,3,debl,Uhlíř Vít,Roušar Petr,W
#
season_year,category,team_label,kolo,datum,souper,pozice,typ,jmeno,partner,vysledek
"""


def init_template(path: str):
    p = Path(path)
    if p.exists():
        print(f"{path} already exists — not overwriting. Delete it first if you want a fresh template.")
        return
    p.write_text(TEMPLATE, encoding="utf-8")
    print(f"Created {path} — fill it in with your match's results, then run:\n"
          f"  python3 add_manual_round.py")


def _detect_dialect(data_text: str) -> csv.Dialect:
    """
    Auto-detect the CSV delimiter instead of assuming comma.
    Czech-locale Excel/LibreOffice export CSV with SEMICOLONS (since comma
    is the Czech decimal separator) -- a file edited and re-saved through
    such a program will silently break a comma-only parser: every line
    becomes one unparsed field, no column is ever found, and the script
    looks like it's "not seeing" data that is clearly there.
    Falls back to comma if there isn't enough signal to sniff confidently.
    """
    try:
        return csv.Sniffer().sniff(data_text, delimiters=",;\t")
    except csv.Error:
        class _Fallback(csv.excel):
            delimiter = ","
        return _Fallback


def read_manual_csv(path: str) -> list[dict]:
    p = Path(path)
    if not p.exists():
        print(f"ERROR: {path} not found. Run with --init first to create a template.")
        sys.exit(1)

    # utf-8-sig strips a UTF-8 BOM if a spreadsheet program added one on
    # save; behaves exactly like plain utf-8 when there's no BOM, so this
    # is a strict improvement either way.
    with open(p, encoding="utf-8-sig") as f:
        raw_lines = f.readlines()

    data_lines = [line for line in raw_lines if not line.lstrip().startswith("#")]
    if not data_lines:
        print(f"ERROR: {path} has no non-comment lines at all (everything starts with '#'). "
              f"Did the file actually save? Try: cat {path}")
        sys.exit(1)

    dialect = _detect_dialect("".join(data_lines))
    delim_name = {",": "comma", ";": "semicolon", "\t": "tab"}.get(dialect.delimiter, repr(dialect.delimiter))
    print(f"  (detected '{delim_name}' as the CSV delimiter)")

    reader = csv.DictReader(data_lines, dialect=dialect)
    fieldnames = reader.fieldnames or []
    missing_cols = [c for c in MANUAL_CSV_HEADER if c not in fieldnames]
    if missing_cols:
        print(f"ERROR: {path}'s header doesn't have the expected columns: {missing_cols}")
        print(f"  Header found: {fieldnames}")
        print(f"  Expected:     {MANUAL_CSV_HEADER}")
        print(f"  If you edited this in Excel/LibreOffice and the columns above look "
              f"merged into one or scrambled, the file may have been saved with a "
              f"different delimiter/encoding than expected -- try re-saving as "
              f"'CSV UTF-8 (Comma delimited)' explicitly.")
        sys.exit(1)

    rows = []
    skipped_blank = 0
    for i, row in enumerate(reader, start=2):
        if not (row.get("jmeno") or "").strip():
            skipped_blank += 1
            continue   # blank/template line

        for key in MANUAL_CSV_HEADER:
            if not (row.get(key) or "").strip() and key not in ("category", "team_label", "partner"):
                print(f"ERROR ({path} line {i}): missing required field '{key}'")
                sys.exit(1)

        typ = row["typ"].strip().lower()
        if typ not in ("singl", "debl"):
            print(f"ERROR ({path} line {i}): typ must be 'singl' or 'debl', got {row['typ']!r}")
            sys.exit(1)

        vysledek = row["vysledek"].strip().upper()
        if vysledek not in ("W", "L"):
            print(f"ERROR ({path} line {i}): vysledek must be 'W' or 'L', got {row['vysledek']!r}")
            sys.exit(1)

        rows.append({
            "season_year": row["season_year"].strip(),
            "category":    (row.get("category") or "").strip() or DEFAULT_CATEGORY,
            "team_label":  (row.get("team_label") or "").strip(),
            "kolo":        row["kolo"].strip(),
            "date":        row["datum"].strip(),
            "opponent":    row["souper"].strip(),
            "position":    row["pozice"].strip(),
            "type":        typ,
            "player_name": row["jmeno"].strip(),
            "partner":     (row.get("partner") or "").strip() or None,
            "result":      vysledek,
        })

    if not rows:
        print(f"ERROR: {path} parsed cleanly ({skipped_blank} blank/template line(s) "
              f"skipped) but found ZERO usable data rows.")
        print(f"  Header found: {fieldnames}")
        print(f"  Most likely cause: the file wasn't actually saved after editing, or "
              f"was saved to a different location than {path}.")
        print(f"  Try: cat {path}   (to see exactly what's on disk right now)")
        sys.exit(1)

    print(f"  Read {len(rows)} data row(s) from {path} "
          f"({skipped_blank} blank/comment line(s) skipped)")
    return rows


def merge_manual_rows(data: dict, manual_rows: list[dict], label: str) -> bool:
    """
    Mutates `data` (one team's bundle) in place: adds the manual round to
    rounds/long_log, then recomputes summary + wide grid.
    Returns True if anything was actually merged.
    """
    if not manual_rows:
        return False

    kolo = manual_rows[0]["kolo"]

    # IMPORTANT: data["rounds"] comes from the team's schedule table, which
    # lists EVERY round for the season including ones not yet played --
    # those show up with zapas_id=None and contribute zero long_log rows.
    # So "this kolo is in data['rounds']" does NOT mean it has real results;
    # only data["long_log"] reflects actual played/recorded matches. Check
    # long_log, not rounds, or every not-yet-published round gets wrongly
    # treated as a duplicate.
    kolos_with_results = {row["kolo"] for row in data["long_log"]}
    if kolo in kolos_with_results:
        print(f"  [{label}] Round {kolo} already has real results from the website scrape — "
              f"SKIPPING manual entry to avoid double-counting. "
              f"(The site has published it since you started filling this in — "
              f"nothing to do here, it's already correct.)")
        return False

    # Resolve each manual player name to the SAME player_id used elsewhere
    # for that person, so their manual-round stats merge into one row
    # instead of creating a duplicate "ghost" entry. Manual rows have no
    # site-assigned id; without this they'd be keyed by name alone while
    # their scraped rows are keyed by id -- two separate people as far as
    # compute_player_summary is concerned, even though it's the same person.
    name_to_id: dict[str, str] = {}
    for p in data["roster"]:
        if p.get("player_id"):
            name_to_id[p["name"]] = p["player_id"]
    for row in data["long_log"]:
        if row["player_id"] and row["player_name"] not in name_to_id:
            name_to_id[row["player_name"]] = row["player_id"]

    unresolved = sorted({
        n for n in {r["player_name"] for r in manual_rows} | {r["partner"] for r in manual_rows if r["partner"]}
        if n not in name_to_id
    })
    if unresolved:
        print(f"  [{label}] WARNING: couldn't match these manual names to an existing "
              f"player_id (check spelling/diacritics against the roster exactly) -- "
              f"they'll be added as BRAND-NEW player entries instead of merging into "
              f"their existing stats:")
        for n in unresolved:
            print(f"      '{n}'")

    # 1. Add the round header for Match Grid columns -- but only if this
    #    kolo isn't ALREADY in the schedule (it usually is: the team page
    #    lists the full season schedule upfront, including future/unplayed
    #    rounds with zapas_id=None). Appending a second entry for the same
    #    kolo would give the Match Grid two columns for one round.
    scheduled_kolos = {r["kolo"] for r in data["rounds"]}
    if kolo not in scheduled_kolos:
        data["rounds"].append({
            "zapas_id": None,
            "kolo": kolo,
            "date": manual_rows[0]["date"],
            "opponent": manual_rows[0]["opponent"],
        })
        data["rounds"].sort(key=lambda r: int(r["kolo"]))

    # 2. Add the long_log rows (player_id resolved above where possible --
    #    falls back to None only for genuinely unrecognised names)
    for r in manual_rows:
        data["long_log"].append({
            "kolo": r["kolo"],
            "date": r["date"],
            "opponent": r["opponent"],
            "position": r["position"],
            "type": r["type"],
            "player_name": r["player_name"],
            "player_id": name_to_id.get(r["player_name"]),
            "result": r["result"],
            "partner": r["partner"],
        })

    # 3. Recompute everything downstream of long_log
    data["summary"] = order_by_roster(compute_player_summary(data["long_log"]), data["roster"])
    data["wide"] = build_wide_grid(data["rounds"], data["long_log"])

    print(f"  [{label}] Merged round {kolo}: {len(manual_rows)} row(s) "
          f"({sum(1 for r in manual_rows if r['result']=='W')} W / "
          f"{sum(1 for r in manual_rows if r['result']=='L')} L)")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Add a manually-entered round to a season and regenerate everything.")
    parser.add_argument("--file", default=MANUAL_CSV_PATH,
                        help=f"Manual round CSV path (default: {MANUAL_CSV_PATH})")
    parser.add_argument("--init", action="store_true",
                        help="Create a template manual_round.csv and exit")
    args = parser.parse_args()

    if args.init:
        init_template(args.file)
        return

    manual_rows = read_manual_csv(args.file)
    if not manual_rows:
        print(f"{args.file} has no data rows yet — fill it in first.")
        return

    # Group manual rows by (season_year, category, team_label)
    manual_by_team: dict[tuple, list[dict]] = {}
    for r in manual_rows:
        key = (r["season_year"], r["category"], r["team_label"])
        manual_by_team.setdefault(key, []).append(r)

    affected_groups = {(s, c) for (s, c, _t) in manual_by_team}
    print(f"Manual entries found for: "
          f"{', '.join(f'{s}/{c} (tým {t or chr(34)+chr(34)})' for s, c, t in manual_by_team)}\n")

    entries = load_teams(TEAMS_CONFIG_PATH)
    groups = group_teams(entries)

    dospeli_seasons_processed = []

    for (season_year, category) in sorted(affected_groups):
        if (season_year, category) not in groups:
            print(f"WARNING: {season_year}/{category} not found in {TEAMS_CONFIG_PATH} — skipping.")
            continue

        print(f"{'=' * 60}\n{season_year} / {category}  (re-scraping all teams in this group)\n{'=' * 60}")
        group_entries = groups[(season_year, category)]
        team_bundles = []
        for entry in group_entries:
            team_label = entry["team_label"].strip()
            try:
                data = process_one_team(entry)
            except Exception as e:
                print(f"  FAILED to scrape {team_label}: {e} — this team will be MISSING "
                      f"from the rewritten workbook, including the All Club view.")
                continue

            key = (season_year, category, team_label)
            if key in manual_by_team:
                merge_manual_rows(data, manual_by_team[key], team_label or "(no label)")

            team_bundles.append((team_label, data))
            print()

        if not team_bundles:
            print(f"  No teams succeeded for {season_year}/{category} — nothing written.\n")
            continue

        out_path = f"{OUTPUT_DIR}/{season_year}_{category}.xlsx"
        write_season_workbook(out_path, team_bundles)
        print(f"  Wrote: {out_path} ({len(team_bundles)} team(s): "
              f"{', '.join(t for t, _ in team_bundles)})\n")

        if category == DEFAULT_CATEGORY:
            dospeli_seasons_processed.append((season_year, team_bundles))

    if dospeli_seasons_processed:
        print("=" * 60)
        print("HISTORY UPDATE")
        for season_year, team_bundles in dospeli_seasons_processed:
            print(f"  Updating history store for {season_year} ...")
            update_history_store(season_year, team_bundles)
        hist_path = f"{OUTPUT_DIR}/club_history_dospeli.xlsx"
        write_history_workbook(hist_path)
        print(f"  Wrote: {hist_path}")

    print()
    print("Done. Charts/HTML weren't touched — re-run these if you want them updated:")
    print("  python3 visualize.py")
    print("  python3 history_charts.py")
    print("  python3 season_awards.py <season>")


if __name__ == "__main__":
    main()
