"""
Aggregation logic: turn parsed round + match-record data into
(1) a long-format match log (one row per player per position played)
(2) a wide grid matching the Excel visual layout (player x round x singl/debl)
(3) a per-player summary table matching the Excel's stat columns exactly:
    Odehrano singl/debl/celkem, Vyhrano singl/debl, Uspesnost singl/debl,
    Body Suma, Body Vazeno (debl win = 0.5 pt)
"""


def build_long_log(
    rounds: list[dict],
    match_records: dict[str, dict],
    name_overrides: dict | None = None,
) -> list[dict]:
    """
    rounds: output of parse_rounds (each has zapas_id, date, opponent, kolo)
    match_records: zapas_id -> output of parse_match_record
    name_overrides: optional {player_id: display_name} from name_overrides.csv;
        applied here so every downstream sheet (Player Summary, Match Grid,
        Match Log, Awards, All Club) sees the canonical name from the start.
    Returns one row per (round, position, our_player).
    """
    overrides = name_overrides or {}
    log = []
    for rnd in rounds:
        zid = rnd["zapas_id"]
        record = match_records.get(zid)
        if record is None:
            continue
        for pos in record["positions"]:
            for player in pos["our_players"]:
                pid = player["player_id"]
                display_name = (
                    overrides.get(str(pid)) if pid else None
                ) or player["name"]
                log.append({
                    "kolo": rnd["kolo"],
                    "date": rnd["date"],
                    "opponent": rnd["opponent"],
                    "position": pos["position"],
                    "type": pos["type"],
                    "player_name": display_name,
                    "player_id": pid,
                    "result": pos["result"],
                    "partner": next(
                        (
                            overrides.get(str(p["player_id"])) or p["name"]
                            for p in pos["our_players"]
                            if p["name"] != player["name"]
                        ),
                        None,
                    ),
                })
    return log


def compute_player_summary(long_log: list[dict]) -> list[dict]:
    """
    One row per player with the same stat columns as the existing Excel.
    Keyed by player_id (falling back to name if id is somehow missing),
    not by name -- this matters once results from multiple teams get
    merged together (the all-club view), where keying by name risks
    silently splitting one person into two rows if they were ever
    entered under slightly different name text in different places.
    Keying by the site's own numeric ID sidesteps that entirely.
    """
    players = {}
    for row in long_log:
        key = row["player_id"] or row["player_name"]
        if key not in players:
            players[key] = {
                "player_name": row["player_name"],
                "player_id": row["player_id"],
                "odehrano_singl": 0, "odehrano_debl": 0,
                "vyhrano_singl": 0, "vyhrano_debl": 0,
            }
        p = players[key]
        if row["result"] is None:
            continue
        stat_key = "singl" if row["type"] == "singl" else "debl"
        p[f"odehrano_{stat_key}"] += 1
        if row["result"] == "W":
            p[f"vyhrano_{stat_key}"] += 1

    summary = []
    for p in players.values():
        os_, od_ = p["odehrano_singl"], p["odehrano_debl"]
        vs_, vd_ = p["vyhrano_singl"], p["vyhrano_debl"]
        summary.append({
            "player_name": p["player_name"],
            "player_id": p["player_id"],
            "odehrano_singl": os_,
            "odehrano_debl": od_,
            "odehrano_celkem": os_ + od_,
            "vyhrano_singl": vs_,
            "vyhrano_debl": vd_,
            "uspesnost_singl": round(vs_ / os_, 4) if os_ else None,
            "uspesnost_debl": round(vd_ / od_, 4) if od_ else None,
            "body_suma": vs_ + vd_,
            "body_vazeno": vs_ * 1.0 + vd_ * 0.5,
        })
    return summary  # order = order of first appearance; call order_by_roster() to reorder


def order_by_roster(summary: list[dict], roster: list[dict]) -> list[dict]:
    """
    Reorder a player summary to match the official Soupiska (roster) order
    from the team page, rather than performance. Roster entries are
    matched by player_id (robust to name-format differences). Any player
    who appears in the data but not in the supplied roster -- e.g. a
    substitute who isn't on the season's published Soupiska -- is placed
    at the end, alphabetically, with a note printed so it's easy to spot.
    """
    roster_position = {}
    for entry in roster:
        pid = entry.get("player_id")
        if pid is None:
            continue
        try:
            roster_position[pid] = int(entry["roster_no"])
        except (TypeError, ValueError):
            continue

    unmatched = [s["player_name"] for s in summary if s["player_id"] not in roster_position]
    if unmatched:
        print(f"  NOTE: {len(unmatched)} player(s) not found on the roster, placed at the end: {', '.join(unmatched)}")

    for s in summary:
        s["roster_no"] = roster_position.get(s["player_id"])  # None if not on the fetched roster

    def sort_key(s):
        pos = s["roster_no"]
        return (1, s["player_name"]) if pos is None else (0, pos)

    return sorted(summary, key=sort_key)


# label, summary field, display kind, conditional-formatting treatment --
# mirrors the "Oceneni" block in the existing Excel exactly (same six
# categories, same ranking metric, and the same color/data-bar choice
# the original file uses for that metric).
AWARD_CATEGORIES = [
    ("MVP", "body_vazeno", "number", "databar_blue"),
    ("Singlový specialista", "uspesnost_singl", "percent", "colorscale"),
    ("Deblový specialista", "uspesnost_debl", "percent", "colorscale"),
    ("Nejvíce odehráno", "odehrano_celkem", "integer", None),
    ("Singl výhry", "vyhrano_singl", "integer", "databar_green"),
    ("Debl výhry", "vyhrano_debl", "integer", "databar_orange"),
]

# Separate category set for the career history Awards sheet.
# Uses Body Suma as the lead "club contribution" metric and adds
# per-season averages + seasons count, which aren't in per-season awards.
HISTORY_AWARD_CATEGORIES = [
    ("Celkem bodů pro klub", "body_suma",       "integer", "databar_steelbl"),
    ("MVP (Váženo)",         "body_vazeno",     "number",  "databar_blue"),
    ("Singlový specialista", "uspesnost_singl", "percent", "colorscale"),
    ("Deblový specialista",  "uspesnost_debl",  "percent", "colorscale"),
    ("Nejvíce odehráno",     "odehrano_celkem", "integer", "databar_magenta"),
    ("Odehráno sezon",       "sezony",          "integer", None),
    ("Singl výhry",          "vyhrano_singl",   "integer", "databar_green"),
    ("Debl výhry",           "vyhrano_debl",    "integer", "databar_orange"),
    ("Ø singl výhry/sezonu", "avg_singl",       "number",  "databar_green"),
    ("Ø debl výhry/sezonu",  "avg_debl",        "number",  "databar_orange"),
    ("Ø váženo/sezonu",      "avg_vazeno",      "number",  "databar_blue"),
]


def compute_oceneni(summary: list[dict], categories_def=None) -> dict:
    """
    Replicates the 'Oceneni' awards table: for each category, every player
    ranked by that category's metric descending. Ties keep their existing
    order (stable sort), so pass a roster-ordered summary for per-team use.
    categories_def defaults to AWARD_CATEGORIES; pass HISTORY_AWARD_CATEGORIES
    for the career history sheet (which has extra fields like sezony, avg_*).
    None values are treated as 0 so every player appears in every ranking.
    """
    cats = categories_def if categories_def is not None else AWARD_CATEGORIES
    result = {}
    for label, field, _kind, _cf in cats:
        ranked = sorted(
            summary,
            key=lambda s: s.get(field) if s.get(field) is not None else 0,
            reverse=True,
        )
        result[label] = [
            {"name": s["player_name"], "value": s[field] if s[field] is not None else 0}
            for s in ranked
        ]
    return result


def build_wide_grid(rounds: list[dict], long_log: list[dict]) -> dict:
    """
    Wide grid for visual side-by-side comparison with the Excel sheet:
    {player_name: {kolo: {"singl": "W"/"L"/None, "debl": "W"/"L"/None}}}
    plus the ordered list of kolo labels (with date + opponent) for headers.
    """
    grid: dict[str, dict] = {}
    for row in long_log:
        name = row["player_name"]
        grid.setdefault(name, {})
        grid[name].setdefault(row["kolo"], {"singl": None, "debl": None})
        grid[name][row["kolo"]][row["type"]] = row["result"]

    headers = [{"kolo": r["kolo"], "date": r["date"], "opponent": r["opponent"]} for r in rounds]
    return {"grid": grid, "headers": headers}


def merge_long_logs(team_bundles: list[tuple[str, dict]]) -> tuple[list[dict], dict[str, list[str]]]:
    """
    Combines every team's long_log into one, for the all-club view of a
    season (players who played for more than one team get their results
    from every team they appeared on summed together once this feeds
    into compute_player_summary).

    team_bundles: list of (team_label, data) where data["long_log"] is
    what process_one_team produced for that team.

    Returns (combined_long_log, teams_by_player_id) where teams_by_player_id
    maps player_id -> sorted list of team_labels that player appeared on
    -- used to show e.g. "A, C" next to a player who played both.
    """
    combined = []
    teams_by_player_id: dict[str, set] = {}
    for team_label, data in team_bundles:
        for row in data["long_log"]:
            tagged = dict(row)
            tagged["team"] = team_label
            combined.append(tagged)
            pid = row["player_id"]
            if pid is not None:
                teams_by_player_id.setdefault(pid, set()).add(team_label)
    return combined, {pid: sorted(teams) for pid, teams in teams_by_player_id.items()}


def compute_all_club_summary(team_bundles: list[tuple[str, dict]]) -> list[dict]:
    """
    The cross-team view: every player who appeared for any team in this
    season+category, with their results from every team they played for
    summed together. Built by merging long_logs (see merge_long_logs)
    and running the result through the same compute_player_summary used
    for a single team -- one aggregation code path, not two, so there's
    no separate "combine the summaries" logic that could drift out of
    sync with how a single team's stats are computed.

    Ordered by Body Vazeno descending, since there's no single roster
    that spans multiple teams to order by instead.
    """
    combined_long_log, teams_by_player_id = merge_long_logs(team_bundles)
    summary = compute_player_summary(combined_long_log)
    for s in summary:
        s["teams"] = "/".join(teams_by_player_id.get(s["player_id"], []))
    summary.sort(key=lambda s: s["body_vazeno"], reverse=True)
    return summary


# ---- History store (persistent JSON accumulator) -------------------------

import json as _json
from pathlib import Path as _Path

HISTORY_STORE_PATH = "history_store.json"
LEGACY_LABEL = "do 2025"   # display label for the through-2025 baseline block


def _load_store(path: str) -> dict:
    p = _Path(path)
    if p.exists():
        try:
            with open(p, encoding="utf-8") as f:
                return _json.load(f)
        except _json.JSONDecodeError as e:
            raise RuntimeError(
                f"\n\nCould not parse {p} as JSON: {e}\n\n"
                f"Most likely cause: the file is corrupt, truncated, or is an older\n"
                f"version from a previous run.  Fix: regenerate it by running:\n\n"
                f"    python3 import_history.py\n\n"
                f"(make sure EXCEL_PATH at the top of that script points to your\n"
                f"current Statistika_mistraky.xlsx)"
            ) from e
    return {"_meta": {}, "players": {}}


def _save_store(store: dict, path: str):
    with open(path, "w", encoding="utf-8") as f:
        _json.dump(store, f, ensure_ascii=False, indent=2)


NAME_OVERRIDES_PATH = "name_overrides.csv"


def _load_name_overrides(path: str = NAME_OVERRIDES_PATH) -> dict[str, str]:
    """
    Loads player_id → canonical_display_name mappings from name_overrides.csv.
    Used to correctly match players whose website name differs from their
    store name (e.g. 'Hýbl Jan' → 'Hýbl Jan st.' or 'Hýbl Jan ml.').
    Returns empty dict if the file doesn't exist or has no valid rows yet.
    """
    overrides: dict[str, str] = {}
    p = _Path(path)
    if not p.exists():
        return overrides
    import csv as _csv
    with open(p, encoding="utf-8") as f:
        for row in _csv.DictReader(f):
            pid = (row.get("player_id") or "").strip()
            name = (row.get("display_name") or "").strip()
            if pid and name and not pid.startswith("#") and not pid.startswith("FILL_IN"):
                overrides[pid] = name
    return overrides


NAME_OVERRIDES_PATH = "name_overrides.csv"


def _load_name_overrides(path: str = NAME_OVERRIDES_PATH) -> dict:
    """
    Load player_id → canonical_name from name_overrides.csv.
    Lines starting with '#' are treated as comments and skipped.
    Placeholder IDs ('FILL_IN_...') are also ignored.
    Returns empty dict if the file doesn't exist.
    """
    import csv as _csv
    overrides: dict[str, str] = {}
    p = _Path(path)
    if not p.exists():
        return overrides
    with open(p, encoding="utf-8") as f:
        # Filter out comment lines before handing to DictReader
        data_lines = (line for line in f if not line.lstrip().startswith("#"))
        for row in _csv.DictReader(data_lines):
            pid  = row.get("player_id",    "").strip()
            name = row.get("display_name", "").strip()
            if pid and name and not pid.startswith("FILL_IN"):
                overrides[pid] = name
    return overrides


def update_history_store(
    season_year: str,
    team_bundles: list[tuple[str, dict]],
    store_path: str = HISTORY_STORE_PATH,
):
    """
    Called by main.py after every dospeli season is processed.
    For each player in the combined all-club summary, writes (or overwrites)
    their stats for `season_year` in the history store.  Running this twice
    for the same season (e.g. mid-season then end-of-season) is safe --
    the second run simply replaces the first.

    Player matching priority:
      1. name_overrides.csv  (player_id → canonical name, handles
         disambiguation like Hýbl Jan st. vs Hýbl Jan ml.)
      2. player_id lookup in store
      3. exact scraped name match
      4. new entry (genuinely new player)
    """
    store = _load_store(store_path)
    players = store.setdefault("players", {})
    overrides = _load_name_overrides()  # player_id → canonical store name
    if overrides:
        print(f"  History: loaded {len(overrides)} name override(s)")

    # Build lookup: player_id → store key name (for fast matching)
    id_to_store_name: dict[str, str] = {}
    for name, entry in players.items():
        pid = entry.get("player_id")
        if pid:
            id_to_store_name[pid] = name

    combined_log, teams_by_id = merge_long_logs(team_bundles)
    season_summary = compute_player_summary(combined_log)

    for s in season_summary:
        pid = s.get("player_id")
        scraped_name = s["player_name"]
        season_teams = "/".join(teams_by_id.get(pid, []))

        # 1. name_overrides.csv: player_id → canonical name
        if pid and pid in overrides:
            store_name = overrides[pid]
            if store_name not in players:
                print(f"  History: WARNING override maps {pid} → '{store_name}' "
                      f"but that name isn't in the store yet")
                continue
            # Backfill player_id in store entry if not set
            if not players[store_name].get("player_id"):
                players[store_name]["player_id"] = pid
                id_to_store_name[pid] = store_name
                print(f"  History: linked player_id {pid} to '{store_name}' via override")

        # 2. player_id lookup in store
        elif pid and pid in id_to_store_name:
            store_name = id_to_store_name[pid]

        # 3. exact scraped name match
        elif scraped_name in players:
            store_name = scraped_name
            # Backfill player_id if we now know it
            if pid and not players[store_name].get("player_id"):
                players[store_name]["player_id"] = pid
                id_to_store_name[pid] = store_name
                print(f"  History: linked player_id {pid} to existing entry '{store_name}'")

        # 4. genuinely new player
        else:
            store_name = scraped_name
            players[store_name] = {
                "nickname":  scraped_name,
                "player_id": pid,
                "teams":     "",
                "_order":    max((e.get("_order", 0) for e in players.values()), default=0) + 1,
                "_legacy": {"os": 0, "od": 0, "vs": 0, "vd": 0, "seasons": 0},
                "seasons": {},
            }
            if pid:
                id_to_store_name[pid] = store_name
            print(f"  History: added new player '{store_name}'")

        # Write / overwrite this season's data
        players[store_name]["seasons"][season_year] = {
            "os": s["odehrano_singl"],
            "od": s["odehrano_debl"],
            "vs": s["vyhrano_singl"],
            "vd": s["vyhrano_debl"],
        }

        # Expand teams field (union of all seasons, never shrinks)
        existing_teams = {
            t.strip()
            for t in players[store_name].get("teams", "").split(",")
            if t.strip()
        }
        new_teams = {t.strip() for t in season_teams.split("/") if t.strip()}
        merged = sorted(existing_teams | new_teams)
        players[store_name]["teams"] = ", ".join(merged)

    _save_store(store, store_path)


def compute_history_rows(store_path: str = HISTORY_STORE_PATH) -> dict:
    """
    Read history_store.json and produce rows for the history sheet.

    Returns:
        {
          "rows": [
            {
              "name": str,
              "teams": str,           # lifetime teams string
              "player_id": str|None,
              "lifetime": {os, od, celkem, vs, vd, us_singl, us_debl, suma, vazeno, seasons},
              "seasons": {
                LEGACY_LABEL: {os, od, vs, vd, vazeno},
                "2024":       {os, od, vs, vd, vazeno},
                "2025":       {os, od, vs, vd, vazeno},
                ...
              }
            }, ...
          ],
          "season_labels": [LEGACY_LABEL, "2024", "2025", ...]  # ordered
        }

    Rows are sorted by lifetime Body Váženo descending.
    Úspěšnost uses the same ≥5-match threshold as the original Excel
    (IF(singl_played < 5, 0, vyhráno/odehráno)).
    """
    store = _load_store(store_path)
    players = store.get("players", {})

    # Collect all scraped year labels in ascending order
    scraped_years: set[str] = set()
    for entry in players.values():
        scraped_years.update(entry.get("seasons", {}).keys())
    season_labels = [LEGACY_LABEL] + sorted(scraped_years)

    rows = []
    for name, entry in players.items():
        leg = entry.get("_legacy", {})
        seasons = entry.get("seasons", {})

        # Lifetime totals (legacy + all scraped seasons)
        os_ = leg.get("os", 0) + sum(s["os"] for s in seasons.values())
        od_ = leg.get("od", 0) + sum(s["od"] for s in seasons.values())
        vs_ = leg.get("vs", 0) + sum(s["vs"] for s in seasons.values())
        vd_ = leg.get("vd", 0) + sum(s["vd"] for s in seasons.values())
        sez = leg.get("seasons", 0) + sum(
            1 for s in seasons.values() if s["os"] + s["od"] > 0
        )

        # Úspěšnost with ≥5 match threshold (matches original Excel formula)
        us_s = (vs_ / os_) if os_ >= 5 else (0.0 if os_ > 0 else None)
        us_d = (vd_ / od_) if od_ >= 5 else (0.0 if od_ > 0 else None)

        # Per-season breakdown: legacy block + one block per scraped year
        season_data = {}
        leg_vaz = leg.get("vs", 0) + 0.5 * leg.get("vd", 0)
        season_data[LEGACY_LABEL] = {
            "os":     leg.get("os", 0),
            "od":     leg.get("od", 0),
            "vs":     leg.get("vs", 0),
            "vd":     leg.get("vd", 0),
            "vazeno": leg_vaz,
            "seasons": leg.get("seasons", 0),
        }
        for year in sorted(scraped_years):
            s = seasons.get(year, {"os": 0, "od": 0, "vs": 0, "vd": 0})
            season_data[year] = {
                "os":    s["os"],
                "od":    s["od"],
                "vs":    s["vs"],
                "vd":    s["vd"],
                "vazeno": s["vs"] + 0.5 * s["vd"],
            }

        rows.append({
            "name":      name,
            "nickname":  entry.get("nickname", ""),
            "teams":     entry.get("teams", ""),
            "player_id": entry.get("player_id"),
            "_order":    entry.get("_order", 999),
            "lifetime": {
                "os": os_, "od": od_, "celkem": os_ + od_,
                "vs": vs_, "vd": vd_,
                "us_singl": us_s, "us_debl": us_d,
                "suma": vs_ + vd_, "vazeno": vs_ + 0.5 * vd_,
                "seasons": sez,
            },
            "seasons": season_data,
        })

    # Player Summary sheet preserves the user's table order (_order).
    # The Awards and per-season sheets sort by Body Váženo instead.
    rows.sort(key=lambda r: r["_order"])
    return {"rows": rows, "season_labels": season_labels}
