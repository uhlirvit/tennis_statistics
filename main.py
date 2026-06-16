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

import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

from fetch import fetch
from parser import parse_roster, parse_rounds, parse_match_record
from aggregate import build_long_log, compute_player_summary, build_wide_grid

# ---------------------------------------------------------------- CONFIG
BASE_URL = "https://cesky-tenis.cz"
COMPETITION_ID = "9029"      # Prazska divize 2025, from the URL you gave
TEAM_NUMBER = "8"            # TK Sport Kolovraty's druzstvoCislo in this competition
CLUB_NAME = "TK Sport Kolovraty"
SEASON_LABEL = "2025_A_tym"
OUTPUT_DIR = "output"
# --------------------------------------------------------------------


def team_url():
    return f"{BASE_URL}/soutez/{COMPETITION_ID}?druzstvoCislo={TEAM_NUMBER}"


def match_url(zapas_id):
    return f"{BASE_URL}/oficialni-zapis/{COMPETITION_ID}?zapas={zapas_id}"


def run():
    print(f"Fetching team page: {team_url()}")
    comp_html = fetch(team_url())
    roster = parse_roster(comp_html)
    rounds = parse_rounds(comp_html, CLUB_NAME)
    print(f"  Found {len(roster)} roster entries, {len(rounds)} rounds.")

    match_records = {}
    for rnd in rounds:
        zid = rnd["zapas_id"]
        if zid is None:
            print(f"  Round {rnd['kolo']}: no zapas id found, skipping.")
            continue
        url = match_url(zid)
        print(f"  Fetching round {rnd['kolo']} ({rnd['date']} vs {rnd['opponent']}): {url}")
        try:
            match_html = fetch(url)
            match_records[zid] = parse_match_record(match_html, CLUB_NAME)
        except RuntimeError as e:
            print(f"    WARNING: {e}")

    long_log = build_long_log(rounds, match_records)
    summary = compute_player_summary(long_log)
    wide = build_wide_grid(rounds, long_log)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f"{SEASON_LABEL}_extracted.xlsx")
    write_workbook(out_path, roster, rounds, long_log, summary, wide)
    print(f"\nDone. Wrote: {out_path}")


# ---------------------------------------------------------------- OUTPUT

HEADER_FILL = PatternFill("solid", start_color="DDEBF7")
HEADER_FONT = Font(bold=True)


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
    headers = ["Player", "Odehrano singl", "Odehrano debl", "Odehrano celkem",
               "Vyhrano singl", "Vyhrano debl", "Uspesnost singl", "Uspesnost debl",
               "Body Suma", "Body Vazeno"]
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    grid_row_for_player = {}  # player_name -> row index on the Match Grid sheet, filled below
    for i, s in enumerate(summary):
        row_idx = i + 2
        ws.append([
            s["player_name"], s["odehrano_singl"], s["odehrano_debl"], s["odehrano_celkem"],
            s["vyhrano_singl"], s["vyhrano_debl"],
            s["uspesnost_singl"], s["uspesnost_debl"],
            s["body_suma"], s["body_vazeno"],
        ])
        ws.cell(row=row_idx, column=7).number_format = "0.0%"
        ws.cell(row=row_idx, column=8).number_format = "0.0%"
    ws.column_dimensions["A"].width = 26
    for col in "BCDEFGHIJ":
        ws.column_dimensions[col].width = 13

    # ---- Sheet 2: Match Grid (visual layout like your Excel) -------
    ws2 = wb.create_sheet("Match Grid")
    headers_info = wide["headers"]
    ws2.cell(row=1, column=1, value="Player")
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
    log_headers = ["Kolo", "Date", "Opponent", "Position", "Type", "Player", "Partner", "Result"]
    ws3.append(log_headers)
    _style_header_row(ws3, 1, len(log_headers))
    for row in long_log:
        ws3.append([
            row["kolo"], row["date"], row["opponent"], row["position"],
            row["type"], row["player_name"], row["partner"] or "", row["result"] or "",
        ])
    for i, w in enumerate([6, 12, 32, 9, 7, 26, 26, 8]):
        ws3.column_dimensions[get_column_letter(i + 1)].width = w

    wb.save(path)


if __name__ == "__main__":
    run()
