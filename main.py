"""
End-to-end scraper for one team / one season on cesky-tenis.cz.

WHAT THIS DOES
  1. Fetches the competition/team page -> gets roster + list of 7 rounds
  2. Fetches each round's official match record -> per-position W/L
  3. Aggregates into a player summary (Odehrano, Vyhrano, Uspesnost, Suma, Vazeno)
  4. Writes everything to an .xlsx with 3 sheets so you can compare
     directly against your existing manual sheet.

HOW TO RUN
  1. Install dependencies once:
       pip install beautifulsoup4 openpyxl
  2. Edit the CONFIG block below if needed (it's already set for
     A tym / 2025 / Prazska divize).
  3. Run:
       python3 main.py
  4. Open the resulting .xlsx in /output and compare to your sheet.

NOTE ON ROBOTS.TXT
  cesky-tenis.cz's robots.txt disallows automated tools, which is why
  this must be run from your own machine rather than by Claude directly.
  This script is intentionally slow (2s delay between requests, 9 requests
  total for one team/season) and only reads public results -- the same
  pages you'd open by hand in a browser.
"""

"""
End-to-end scraper for cesky-tenis.cz team competitions.

WHAT THIS DOES
  Reads teams.csv (one row per team-per-season you want tracked).
  For each row:
    1. Fetches the competition/team page -> roster + list of rounds
    2. Fetches each round's official match record -> per-position W/L
    3. Aggregates into a player summary (Odehráno, Vyhráno, Úspěšnost, Suma, Váženo)
    4. Writes one .xlsx per row, named output/{season_year}_{team_label}_extracted.xlsx

HOW TO RUN
  1. Install dependencies once:
       pip install beautifulsoup4 openpyxl
  2. Open teams.csv and make sure every season/team you want is listed
     (see "ADDING A NEW SEASON OR TEAM" below).
  3. Run:
       python3 main.py
  4. Each row produces its own file in /output. Compare against your
     existing sheet for that season/team.

ADDING A NEW SEASON OR TEAM
  Add a row to teams.csv. No code changes needed. You need four things
  per row, all read straight from the team's URL on the website, e.g.
  https://cesky-tenis.cz/soutez/9029?druzstvoCislo=8
                              ^^^^                ^
                       competition_id      team_number
  - season_year: just a label for the filename, e.g. 2024
  - team_label: A / B / C / dorost / etc., also just a filename label
  - competition_name: free text, purely for your own readability
  - competition_id, team_number: the two numbers from the URL above
  If you're not sure which URL to use, find the team's standings page
  on the site the same way you did originally, or paste the URL here
  and it'll get translated into a row directly.

NOTE ON ROBOTS.TXT
  cesky-tenis.cz's robots.txt disallows automated tools, which is why
  this must be run from your own machine rather than by Claude directly.
  This script is intentionally slow (2s delay between requests) and
  only reads public results -- the same pages you'd open by hand in a
  browser. Running N rows fetches roughly 9*N pages in total.
"""

import os
import csv
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import DataBarRule, ColorScaleRule

from fetch import fetch
from parser import parse_roster, parse_rounds, parse_match_record, parse_club_display_name
from aggregate import (
    build_long_log, compute_player_summary, order_by_roster,
    compute_oceneni, AWARD_CATEGORIES, build_wide_grid,
)

# ---------------------------------------------------------------- CONFIG
BASE_URL = "https://cesky-tenis.cz"
CLUB_NAME = "TK Sport Kolovraty"
TEAMS_CONFIG_PATH = "teams.csv"
OUTPUT_DIR = "output"
# --------------------------------------------------------------------


def team_url(competition_id, team_number):
    return f"{BASE_URL}/soutez/{competition_id}?druzstvoCislo={team_number}"


def match_url(competition_id, zapas_id):
    return f"{BASE_URL}/oficialni-zapis/{competition_id}?zapas={zapas_id}"


def parse_team_url(url):
    """
    Pulls competition_id and team_number straight out of a team URL like
    https://cesky-tenis.cz/soutez/9029?druzstvoCislo=8
    so teams.csv only needs to store the URL itself -- one source of
    truth instead of the URL plus two numbers that could drift out of
    sync with it.
    """
    comp_match = re.search(r"/soutez/(\d+)", url)
    team_match = re.search(r"druzstvoCislo=(\d+)", url)
    if not comp_match or not team_match:
        raise ValueError(f"Could not find competition_id/team_number in URL: {url!r}")
    return comp_match.group(1), team_match.group(1)


def load_teams(path):
    with open(path, encoding="utf-8") as f:
        return list(csv.DictReader(f))


def run_one(entry):
    url = entry["url"].strip()
    competition_id, team_number = parse_team_url(url)
    season_year = entry["season_year"].strip()
    team_label = entry["team_label"].strip()
    label = f"{season_year} {team_label} ({entry.get('competition_name', '').strip()})"

    print(f"=== {label} ===")
    print(f"  URL: {url}")
    print(f"  Parsed competition_id={competition_id}, team_number={team_number}")
    print(f"Fetching team page: {team_url(competition_id, team_number)}")
    comp_html = fetch(team_url(competition_id, team_number))

    club_display_name = parse_club_display_name(comp_html)
    if not club_display_name:
        print(f"  WARNING: couldn't read a team name off this page at all -- falling back to {CLUB_NAME!r}")
        club_display_name = CLUB_NAME
    elif CLUB_NAME not in club_display_name:
        print(f"  WARNING: page's team name is {club_display_name!r}, which doesn't contain "
              f"{CLUB_NAME!r} -- double check the URL/team_number for this row, this might be the wrong team.")
    elif club_display_name != CLUB_NAME:
        print(f"  Note: this team's exact name on the site is {club_display_name!r}, using that.")

    roster = parse_roster(comp_html)
    rounds = parse_rounds(comp_html, club_display_name)
    print(f"  Found {len(roster)} roster entries, {len(rounds)} rounds.")

    match_records = {}
    for rnd in rounds:
        zid = rnd["zapas_id"]
        if zid is None:
            print(f"  Round {rnd['kolo']}: no zapas id found, skipping.")
            continue
        url = match_url(competition_id, zid)
        print(f"  Fetching round {rnd['kolo']} ({rnd['date']} vs {rnd['opponent']}): {url}")
        try:
            match_html = fetch(url)
            match_records[zid] = parse_match_record(match_html, club_display_name)
        except RuntimeError as e:
            print(f"    WARNING: {e}")

    long_log = build_long_log(rounds, match_records)
    if match_records and not long_log:
        raise RuntimeError(
            f"Fetched {len(match_records)} match record(s) but extracted 0 player "
            f"results for any of them. Most likely cause: the team's name on the "
            f"site doesn't match CLUB_NAME ({CLUB_NAME!r}) -- check whether this "
            f"division shows it with a suffix, e.g. 'TK Sport Kolovraty C'."
        )
    summary = compute_player_summary(long_log)
    summary = order_by_roster(summary, roster)
    wide = build_wide_grid(rounds, long_log)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f"{season_year}_{team_label}_extracted.xlsx")
    write_workbook(out_path, roster, rounds, long_log, summary, wide)
    print(f"  Done. Wrote: {out_path}\n")
    return out_path


def run_all():
    entries = load_teams(TEAMS_CONFIG_PATH)
    print(f"Loaded {len(entries)} row(s) from {TEAMS_CONFIG_PATH}\n")
    results = []
    for entry in entries:
        try:
            out_path = run_one(entry)
            results.append((entry, "OK", out_path))
        except Exception as e:
            print(f"  FAILED: {e}\n")
            results.append((entry, "FAILED", str(e)))

    print("=" * 60)
    print("SUMMARY")
    for entry, status, detail in results:
        tag = f"{entry['season_year']} {entry['team_label']}"
        print(f"  [{status}] {tag:<12} {detail}")


# ---------------------------------------------------------------- OUTPUT

HEADER_FILL = PatternFill("solid", start_color="DDEBF7")
HEADER_FONT = Font(bold=True)

# Exact colors pulled from the original workbook's conditional formatting
# rules (same metric -> same color in both the main table and Oceneni).
DATABAR_GREEN = "FF63C384"    # Vyhrano singl
DATABAR_ORANGE = "FFFFB628"   # Vyhrano debl
DATABAR_BLUE = "FF008AEF"     # Body Vazeno
SCALE_RED = "FFF8696B"        # Uspesnost -- low end
SCALE_YELLOW = "FFFFEB84"     # Uspesnost -- midpoint (50th percentile)
SCALE_GREEN = "FF63BE7B"      # Uspesnost -- high end

CF_DATABARS = {
    "databar_green": DATABAR_GREEN,
    "databar_orange": DATABAR_ORANGE,
    "databar_blue": DATABAR_BLUE,
}


def _add_databar(ws, col_letter, first_row, last_row, color):
    if last_row < first_row:
        return
    rng = f"{col_letter}{first_row}:{col_letter}{last_row}"
    ws.conditional_formatting.add(
        rng, DataBarRule(start_type="min", start_value=None, end_type="max", end_value=None, color=color)
    )


def _add_colorscale_ryg(ws, col_letter, first_row, last_row):
    if last_row < first_row:
        return
    rng = f"{col_letter}{first_row}:{col_letter}{last_row}"
    ws.conditional_formatting.add(
        rng,
        ColorScaleRule(
            start_type="min", start_color=SCALE_RED,
            mid_type="percentile", mid_value=50, mid_color=SCALE_YELLOW,
            end_type="max", end_color=SCALE_GREEN,
        ),
    )


def _add_cf(ws, cf_kind, col_letter, first_row, last_row):
    if cf_kind is None:
        return
    if cf_kind == "colorscale":
        _add_colorscale_ryg(ws, col_letter, first_row, last_row)
    elif cf_kind in CF_DATABARS:
        _add_databar(ws, col_letter, first_row, last_row, CF_DATABARS[cf_kind])


def _style_header_row(ws, row_idx, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def write_workbook(path, roster, rounds, long_log, summary, wide):
    wb = Workbook()

    # ---- Sheet 1: Player Summary -----------------------------------
    ws = wb.active
    ws.title = "Player Summary"
    headers = ["Č. na soupisce", "Jméno", "Odehráno singl", "Odehráno debl", "Odehráno celkem",
               "Vyhráno singl", "Vyhráno debl", "Úspěšnost singl", "Úspěšnost debl",
               "Body Suma", "Body Váženo"]
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    grid_row_for_player = {}  # player_name -> row index on the Match Grid sheet, filled below
    for i, s in enumerate(summary):
        row_idx = i + 2
        ws.append([
            s["roster_no"], s["player_name"], s["odehrano_singl"], s["odehrano_debl"], s["odehrano_celkem"],
            s["vyhrano_singl"], s["vyhrano_debl"],
            s["uspesnost_singl"], s["uspesnost_debl"],
            s["body_suma"], s["body_vazeno"],
        ])
        ws.cell(row=row_idx, column=8).number_format = "0.0%"
        ws.cell(row=row_idx, column=9).number_format = "0.0%"
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 26
    for col in "CDEFGHIJK":
        ws.column_dimensions[col].width = 13

    last_row = 1 + len(summary)
    _add_databar(ws, "F", 2, last_row, DATABAR_GREEN)      # Vyhrano singl
    _add_databar(ws, "G", 2, last_row, DATABAR_ORANGE)     # Vyhrano debl
    _add_colorscale_ryg(ws, "H", 2, last_row)               # Uspesnost singl
    _add_colorscale_ryg(ws, "I", 2, last_row)               # Uspesnost debl
    _add_databar(ws, "K", 2, last_row, DATABAR_BLUE)        # Body Vazeno

    # ---- Sheet 2: Match Grid (visual layout like your Excel) -------
    ws2 = wb.create_sheet("Match Grid")
    headers_info = wide["headers"]
    ws2.cell(row=1, column=1, value="Jméno")
    col = 2
    for h in headers_info:
        ws2.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        ws2.cell(row=1, column=col, value=f"{h['date']}\n{h['opponent']}")
        ws2.cell(row=2, column=col, value="singl")
        ws2.cell(row=2, column=col + 1, value="debl")
        col += 2
    _style_header_row(ws2, 1, col - 1)
    _style_header_row(ws2, 2, col - 1)
    ws2.row_dimensions[1].height = 40

    players_in_order = [s["player_name"] for s in summary]
    for i, name in enumerate(players_in_order):
        row_idx = i + 3
        ws2.cell(row=row_idx, column=1, value=name)
        grid_row_for_player[name] = row_idx
        col = 2
        for h in headers_info:
            cell_data = wide["grid"].get(name, {}).get(h["kolo"], {"singl": None, "debl": None})
            for j, kind in enumerate(("singl", "debl")):
                val = cell_data[kind]
                cell = ws2.cell(row=row_idx, column=col + j)
                if val == "W":
                    cell.value = 1
                elif val == "L":
                    cell.value = 0
                else:
                    cell.value = None
            col += 2
    ws2.column_dimensions["A"].width = 26

    # ---- Sheet 3: Match Log (audit trail) ---------------------------
    ws3 = wb.create_sheet("Match Log")
    log_headers = ["Kolo", "Datum", "Soupeř", "Pozice", "Typ", "Jméno", "Partner / partnerka", "Výsledek"]
    ws3.append(log_headers)
    _style_header_row(ws3, 1, len(log_headers))
    for row in long_log:
        ws3.append([
            row["kolo"], row["date"], row["opponent"], row["position"],
            row["type"], row["player_name"], row["partner"] or "", row["result"] or "",
        ])
    for i, w in enumerate([6, 12, 32, 9, 7, 26, 26, 10]):
        ws3.column_dimensions[get_column_letter(i + 1)].width = w

    # ---- Sheet 4: Awards / rankings (matches the "Oceneni" block) ---
    ws4 = wb.create_sheet("Awards")
    oceneni = compute_oceneni(summary)
    categories = list(oceneni.keys())
    percent_cats = {label for label, _field, kind, _cf in AWARD_CATEGORIES if kind == "percent"}

    ws4.cell(row=1, column=1, value="Pořadí")
    col = 2
    hodnota_col_for_category = {}
    for cat in categories:
        ws4.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        ws4.cell(row=1, column=col, value=cat)
        hodnota_col_for_category[cat] = col + 1
        col += 2
    _style_header_row(ws4, 1, col - 1)

    col = 2
    for cat in categories:
        ws4.cell(row=2, column=col, value="Jméno")
        ws4.cell(row=2, column=col + 1, value="Hodnota")
        col += 2
    _style_header_row(ws4, 2, col - 1)

    for i in range(len(summary)):
        row_idx = i + 3
        ws4.cell(row=row_idx, column=1, value=f"{i + 1}.")
        col = 2
        for cat in categories:
            entry = oceneni[cat][i]
            ws4.cell(row=row_idx, column=col, value=entry["name"])
            value_cell = ws4.cell(row=row_idx, column=col + 1, value=entry["value"])
            if cat in percent_cats:
                value_cell.number_format = "0.0%"
            col += 2

    ws4.column_dimensions["A"].width = 8
    col = 2
    for _ in categories:
        ws4.column_dimensions[get_column_letter(col)].width = 22
        ws4.column_dimensions[get_column_letter(col + 1)].width = 10
        col += 2

    last_row_oc = 2 + len(summary)
    cf_for_category = {label: cf for label, _field, _kind, cf in AWARD_CATEGORIES}
    for cat in categories:
        cf_kind = cf_for_category.get(cat)
        col_letter = get_column_letter(hodnota_col_for_category[cat])
        _add_cf(ws4, cf_kind, col_letter, 3, last_row_oc)

    wb.save(path)


if __name__ == "__main__":
    run_all()
