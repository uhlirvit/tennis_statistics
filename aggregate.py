"""
Aggregation logic: turn parsed round + match-record data into
(1) a long-format match log (one row per player per position played)
(2) a wide grid matching the Excel visual layout (player x round x singl/debl)
(3) a per-player summary table matching the Excel's stat columns exactly:
    Odehrano singl/debl/celkem, Vyhrano singl/debl, Uspesnost singl/debl,
    Body Suma, Body Vazeno (debl win = 0.5 pt)
"""


def build_long_log(rounds: list[dict], match_records: dict[str, dict]) -> list[dict]:
    """
    rounds: output of parse_rounds (each has zapas_id, date, opponent, kolo)
    match_records: zapas_id -> output of parse_match_record
    Returns one row per (round, position, our_player).
    """
    log = []
    for rnd in rounds:
        zid = rnd["zapas_id"]
        record = match_records.get(zid)
        if record is None:
            continue
        for pos in record["positions"]:
            for player in pos["our_players"]:
                log.append({
                    "kolo": rnd["kolo"],
                    "date": rnd["date"],
                    "opponent": rnd["opponent"],
                    "position": pos["position"],
                    "type": pos["type"],
                    "player_name": player["name"],
                    "player_id": player["player_id"],
                    "result": pos["result"],
                    "partner": next(
                        (p["name"] for p in pos["our_players"] if p["name"] != player["name"]),
                        None,
                    ),
                })
    return log


def compute_player_summary(long_log: list[dict]) -> list[dict]:
    """One row per player with the same stat columns as the existing Excel."""
    players = {}
    for row in long_log:
        name = row["player_name"]
        if name not in players:
            players[name] = {
                "player_name": name,
                "player_id": row["player_id"],
                "odehrano_singl": 0, "odehrano_debl": 0,
                "vyhrano_singl": 0, "vyhrano_debl": 0,
            }
        p = players[name]
        if row["result"] is None:
            continue
        key = "singl" if row["type"] == "singl" else "debl"
        p[f"odehrano_{key}"] += 1
        if row["result"] == "W":
            p[f"vyhrano_{key}"] += 1

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
    summary.sort(key=lambda r: r["body_vazeno"], reverse=True)
    return summary


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
