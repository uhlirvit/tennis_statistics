"""
history_charts.py — Interactive career history charts for TK Sport Kolovraty.

Reads from history_store.json and generates a single self-contained HTML file
with four interactive Plotly charts:

  1. Career leaderboard   — all players, sortable by any metric via dropdown
  2. Career timeline      — approximate career spans; exact for Tříška 2006-2012
  3. Player radar         — normalized 5-metric comparison; top 10 shown by default
  4. Career value         — seasons (experience) vs avg Váženo/season (efficiency)

Usage:
    python3 history_charts.py
    python3 history_charts.py --store path/to/history_store.json
    python3 history_charts.py --out output/my_report.html

Output: output/club_history_interactive.html
Requires: pip install plotly
"""

import argparse
import json
import os
from collections import defaultdict
from pathlib import Path

import plotly.graph_objects as go
import plotly.io as pio

# ── Config ─────────────────────────────────────────────────────────────────
HISTORY_STORE = "history_store.json"
OUTPUT        = "output/club_history_interactive.html"

TEAM_COLORS = {
    "A": "#1A6FBF",
    "B": "#E87722",
    "C": "#2A9D3F",
    "D": "#8E44AD",
    "":  "#AAAAAA",
}
RADAR_PALETTE = [
    "#1f77b4","#ff7f0e","#2ca02c","#d62728","#9467bd",
    "#8c564b","#e377c2","#7f7f7f","#bcbd22","#17becf",
    "#aec7e8","#ffbb78","#98df8a","#ff9896","#c5b0d5",
]

def _team_color(teams_str: str) -> str:
    if not teams_str:
        return TEAM_COLORS[""]
    first = teams_str.split(",")[0].strip()
    return TEAM_COLORS.get(first, TEAM_COLORS[""])


# ── Data loading ────────────────────────────────────────────────────────────

def load_players(store_path: str) -> list[dict]:
    """
    Load history_store.json and compute lifetime totals + derived metrics
    for each player. Works whether or not scraped per-year data is present.
    """
    with open(store_path, encoding="utf-8") as f:
        store = json.load(f)

    players = []
    for name, entry in store["players"].items():
        leg  = entry.get("_legacy", {})
        seas = entry.get("seasons", {})

        os_  = leg.get("os", 0) + sum(s["os"] for s in seas.values())
        od_  = leg.get("od", 0) + sum(s["od"] for s in seas.values())
        vs_  = leg.get("vs", 0) + sum(s["vs"] for s in seas.values())
        vd_  = leg.get("vd", 0) + sum(s["vd"] for s in seas.values())
        sez  = leg.get("seasons", 0) + sum(
            1 for s in seas.values() if s["os"] + s["od"] > 0
        )

        vazeno  = vs_ + 0.5 * vd_
        suma    = vs_ + vd_
        celkem  = os_ + od_
        us_s    = vs_ / os_ if os_ >= 5 else None
        us_d    = vd_ / od_ if od_ >= 5 else None
        avg_vaz = vazeno / sez if sez > 0 else 0.0

        players.append({
            "name":       name,
            "nickname":   entry.get("nickname", name),
            "teams":      entry.get("teams", ""),
            "player_id":  entry.get("player_id"),
            "os": os_, "od": od_, "celkem": celkem,
            "vs": vs_, "vd": vd_,
            "us_singl":   us_s,
            "us_debl":    us_d,
            "suma":       suma,
            "vazeno":     vazeno,
            "seasons":    sez,
            "avg_vazeno": avg_vaz,
            "_legacy":    leg,
            "_seasons":   seas,
            "_order":     entry.get("_order", 999),
        })

    players.sort(key=lambda p: p["_order"])
    return players


# ── Chart 1: Career leaderboard ─────────────────────────────────────────────

def fig_leaderboard(players: list[dict]) -> go.Figure:
    """
    Horizontal bar chart with a dropdown to switch between metrics.
    Each metric creates its own trace; only one is visible at a time.
    """
    metrics = [
        ("vazeno",     "Body Váženo",           ".1f",  False),
        ("suma",       "Body Suma",              "d",    False),
        ("vs",         "Singl výhry",            "d",    False),
        ("vd",         "Debl výhry",             "d",    False),
        ("us_singl",   "Úspěšnost singl",        ".0%",  True),
        ("us_debl",    "Úspěšnost debl",         ".0%",  True),
        ("celkem",     "Odehráno celkem",         "d",    False),
        ("seasons",    "Odehráno sezón",          "d",    False),
        ("avg_vazeno", "Ø Váženo / sezóna",      ".2f",  False),
    ]

    traces  = []
    buttons = []

    for i, (field, label, fmt, is_pct) in enumerate(metrics):
        valid = sorted(
            [p for p in players if p[field] is not None and p[field] > 0],
            key=lambda p: p[field],   # ascending → longest bar at top
        )
        names  = [p["name"] for p in valid]
        values = [p[field] * (100 if is_pct else 1) for p in valid]
        colors = [_team_color(p["teams"]) for p in valid]
        texts  = [
            f"{v:.0f}%" if is_pct else
            (f"{v:.1f}" if isinstance(v, float) and not v.is_integer() else f"{int(v)}")
            for v in values
        ]
        hovers = [
            f"<b>{p['name']}</b><br>"
            f"Tým: {p['teams'] or '—'}<br>"
            f"Sezóny: {p['seasons']}<br>"
            f"Body Váženo: {p['vazeno']:.1f}  |  Suma: {p['suma']}<br>"
            f"Singl: {p['vs']}/{p['os']}  |  Debl: {p['vd']}/{p['od']}"
            for p in valid
        ]

        traces.append(go.Bar(
            x=values, y=names,
            orientation="h",
            marker_color=colors, opacity=0.85,
            text=texts, textposition="outside",
            textfont=dict(size=10),
            hovertext=hovers, hoverinfo="text",
            visible=(i == 0),
            name=label,
            showlegend=False,          # ← metric name must NOT appear in legend
        ))

        # Team legend traces come AFTER all metric traces (4 of them),
        # so the full visibility array is: [metric_flags...] + [True, True, True, True]
        n_team_traces = sum(1 for t in TEAM_COLORS if t)   # = 4
        x_title = f"{label} (%)" if is_pct else label
        buttons.append(dict(
            label=label,
            method="update",
            args=[
                {"visible": [j == i for j in range(len(metrics))]
                            + [True] * n_team_traces},
                {"xaxis.title": x_title,
                 "title.text":  f"Pořadí — {label}"},
            ],
        ))

    # Team colour legend entries — use Scatter markers (invisible data,
    # visible legend square), which work correctly in Plotly 6
    for team, color in sorted(TEAM_COLORS.items()):
        if team:
            traces.append(go.Scatter(
                x=[None], y=[None],
                mode="markers",
                marker=dict(size=14, color=color, symbol="square"),
                name=f"Tým {team}",
                showlegend=True,
            ))

    n_players = max(len([x for x in players if x[f] is not None and x[f] > 0])
                    for f, *_ in metrics)
    fig = go.Figure(data=traces)
    fig.update_layout(
        title="Pořadí — Body Váženo",
        height=max(520, n_players * 26 + 120),
        xaxis_title="Body Váženo",
        yaxis=dict(
            tickfont=dict(size=13, color="#1A1A1A"),  # ← sharper, larger names
            automargin=True,
        ),
        plot_bgcolor="#FAFAFA",
        paper_bgcolor="white",
        margin=dict(l=20, r=120, t=70, b=50),
        legend=dict(
            orientation="v", x=1.01, y=1, xanchor="left",
            title=dict(text="Tým", font=dict(size=12)),
        ),
        updatemenus=[dict(
            type="dropdown",
            direction="down",
            buttons=buttons,
            x=0.0, xanchor="left",
            y=1.06, yanchor="top",
            bgcolor="white",
            bordercolor="#CCCCCC",
            font=dict(size=12),
            showactive=True,
        )],
        barmode="overlay",
    )
    return fig


# ── Chart 2: Career timeline ─────────────────────────────────────────────────

def _career_span(p: dict) -> tuple[int, int]:
    """
    Return (estimated_start, last_active) for a player.
    For most players the last active year is the latest in their scraped
    seasons, or 2025 if only legacy data is present.
    Start is estimated as last_active − (seasons − 1).
    """
    seas = p["_seasons"]
    last = int(max(seas.keys())) if seas else 2025
    start = last - p["seasons"] + 1
    return start, last


def fig_timeline(players: list[dict]) -> go.Figure:
    """
    Gantt-style career timeline showing approximate career spans.
    Tříška Martin gets exact treatment (2006-2012 + 2026 if played).
    Players sorted by estimated career start (oldest at top).
    """
    sorted_p = sorted(players, key=lambda p: (
        2006 if "Tříška" in p["name"] else _career_span(p)[0]
    ))

    fig = go.Figure()

    seen_teams: set[str] = set()

    for p in sorted_p:
        name   = p["name"]
        color  = _team_color(p["teams"])
        teams  = p["teams"] or "Neuveden"
        sleg   = seen_teams
        show_l = teams not in seen_teams
        seen_teams.add(teams)

        hover_base = (
            f"<b>{name}</b><br>"
            f"Přezdívka: {p['nickname']}<br>"
            f"Tým: {teams}<br>"
            f"Celkem sezón: {p['seasons']}<br>"
            f"Body Váženo: {p['vazeno']:.1f}<br>"
            f"Odehráno: {p['celkem']}"
        )

        if "Tříška" in name:
            # Explicit 2006-2012 block
            fig.add_trace(go.Bar(
                y=[name], x=[7], base=[2006],
                orientation="h",
                marker=dict(color=color, opacity=0.80,
                            line=dict(width=0)),
                name=teams, legendgroup=teams,
                showlegend=show_l,
                hovertext=hover_base + "<br><i>2006–2012 (7 sezón)</i>",
                hoverinfo="text",
            ))
            show_l = False
            # 2026 if played (shown as a thin separate block)
            if p["_seasons"].get("2026"):
                s26 = p["_seasons"]["2026"]
                fig.add_trace(go.Bar(
                    y=[name], x=[1], base=[2026],
                    orientation="h",
                    marker=dict(color=color, opacity=0.80,
                                line=dict(width=0)),
                    name=teams, legendgroup=teams,
                    showlegend=False,
                    hovertext=(
                        hover_base +
                        f"<br><i>2026 – návrat!</i><br>"
                        f"Váženo 2026: {s26['vs'] + 0.5 * s26['vd']:.1f}"
                    ),
                    hoverinfo="text",
                ))
        else:
            start, last = _career_span(p)
            note = "" if p["_seasons"] else " <i>(odhadnuto)</i>"
            fig.add_trace(go.Bar(
                y=[name], x=[last - start + 1], base=[start],
                orientation="h",
                marker=dict(color=color,
                            opacity=0.80 if p["_seasons"] else 0.50,
                            line=dict(width=0)),
                name=teams, legendgroup=teams,
                showlegend=show_l,
                hovertext=hover_base + f"<br>~{start}–{last}{note}",
                hoverinfo="text",
            ))

    # Highlight current year
    fig.add_vline(x=2026, line_dash="dot", line_color="#888888",
                  annotation_text="2026", annotation_position="top")

    fig.update_layout(
        title="Kariérní přehled — kdo hrál kdy",
        barmode="overlay",
        height=max(520, len(players) * 24 + 140),
        xaxis=dict(
            title="Rok",
            range=[2003, 2028],
            dtick=1,
            tickangle=-45,
            tickfont=dict(size=10),
            gridcolor="#EEEEEE",
        ),
        yaxis=dict(
            autorange="reversed",
            tickfont=dict(size=13, color="#1A1A1A"),
            automargin=True,
        ),
        plot_bgcolor="#FAFAFA",
        paper_bgcolor="white",
        margin=dict(l=20, r=60, t=70, b=70),
        legend=dict(
            orientation="v",
            title="Tým",
            x=1.01, y=1, xanchor="left",
        ),
        annotations=[dict(
            x=0.0, y=1.04,
            xref="paper", yref="paper",
            text="<i>Šedé / průhledné pruhy = odhadovaný rozsah kariéry.</i>",
            showarrow=False, font=dict(size=10, color="#888888"),
        )],
    )
    return fig


# ── Chart 3: Radar comparison ────────────────────────────────────────────────

def fig_radar(players: list[dict], top_n: int = 10) -> go.Figure:
    """
    Polar radar chart normalized across 5 career metrics.
    Top N by career Váženo shown by default; others toggleable via legend.
    """
    categories = [
        "Singl %", "Debl %", "Ø Váženo / sez",
        "Odehráno", "Sezóny",
    ]

    max_avg = max((p["avg_vazeno"] for p in players if p["avg_vazeno"] > 0), default=1.0)
    max_cel = max((p["celkem"] for p in players), default=1)
    max_sez = max((p["seasons"] for p in players), default=1)

    def normalize(p: dict) -> list[float]:
        return [
            (p["us_singl"] or 0) * 100,
            (p["us_debl"]  or 0) * 100,
            p["avg_vazeno"] / max_avg * 100,
            p["celkem"]     / max_cel * 100,
            p["seasons"]    / max_sez * 100,
        ]

    sorted_p = sorted(players, key=lambda p: -p["vazeno"])

    fig = go.Figure()

    for i, p in enumerate(sorted_p):
        vals   = normalize(p)
        closed = vals + [vals[0]]
        cats   = categories + [categories[0]]
        color  = RADAR_PALETTE[i % len(RADAR_PALETTE)]

        # Build one hover string and repeat it for every point in the
        # closed polygon — customdata[0][...] only works for point 0,
        # causing blank hovers on all other polygon vertices.
        hover_str = (
            f"<b>{p['name']}</b><br>"
            f"Tým: {p['teams'] or '—'}<br>"
            f"Singl %: {(p['us_singl'] or 0)*100:.0f}%<br>"
            f"Debl %:  {(p['us_debl']  or 0)*100:.0f}%<br>"
            f"Ø Váženo/sez: {p['avg_vazeno']:.2f}<br>"
            f"Odehráno: {p['celkem']}<br>"
            f"Sezóny: {p['seasons']}<br>"
            f"Celkem Váženo: {p['vazeno']:.1f}"
        )

        fig.add_trace(go.Scatterpolar(
            r=closed,
            theta=cats,
            fill="toself",
            name=p["name"],
            opacity=0.55,
            line=dict(color=color, width=2),
            fillcolor=color,
            visible=True if i < top_n else "legendonly",
            text=[hover_str] * len(closed),   # same text on every vertex
            hoverinfo="text",
        ))

    fig.update_layout(
        title=(
            f"Kariérní srovnání — radar<br>"
            f"<sup>Top {top_n} zobrazeno; ostatní zapnout/vypnout v legendě</sup>"
        ),
        polar=dict(
            radialaxis=dict(
                visible=True, range=[0, 100],
                tickfont=dict(size=9),
                gridcolor="#DDDDDD",
            ),
            angularaxis=dict(
                tickfont=dict(size=12),
                gridcolor="#DDDDDD",
            ),
            bgcolor="#FAFAFA",
        ),
        showlegend=True,
        legend=dict(
            orientation="v",
            x=1.05, y=1, xanchor="left",
            font=dict(size=10),
        ),
        paper_bgcolor="white",
        height=650,
        margin=dict(l=60, r=220, t=90, b=60),
    )
    return fig


# ── Chart 4: Career value ────────────────────────────────────────────────────

def fig_career_value(players: list[dict]) -> go.Figure:
    """
    X: Seasons (experience)
    Y: Avg Váženo per season (efficiency)
    Size: Total Body Váženo (career impact)
    Color: team
    Labels: last name

    Four natural quadrants emerge:
      top-right   → Legendy (experienced AND efficient)
      top-left    → Specialisté (efficient but fewer seasons)
      bottom-right → Stálice (many seasons, growing)
      bottom-left  → Nováčci
    """
    active = [p for p in players if p["seasons"] > 0 and p["celkem"] > 0]

    by_team: dict[str, list] = defaultdict(list)
    for p in active:
        primary = (p["teams"].split(",")[0].strip() if p["teams"] else "")
        by_team[primary].append(p)

    TEAM_NAMES = {"A": "Tým A", "B": "Tým B", "C": "Tým C", "D": "Tým D", "": "Ostatní"}

    # Medians for quadrant lines
    all_x = [p["seasons"]    for p in active]
    all_y = [p["avg_vazeno"] for p in active]
    med_x = sorted(all_x)[len(all_x) // 2]
    med_y = sorted(all_y)[len(all_y) // 2]

    fig = go.Figure()

    for primary, group in sorted(by_team.items()):
        color = _team_color(primary)
        hovers = [
            f"<b>{p['name']}</b><br>"
            f"Tým: {p['teams'] or '—'}<br>"
            f"Sezóny: {p['seasons']}<br>"
            f"Ø Váženo/sez: {p['avg_vazeno']:.2f}<br>"
            f"Celkem Váženo: {p['vazeno']:.1f}<br>"
            f"Odehráno: {p['celkem']}<br>"
            f"Singl: {p['vs']}/{p['os']}  Debl: {p['vd']}/{p['od']}"
            for p in group
        ]
        sizes = [max(10, p["vazeno"] ** 0.55 * 4) for p in group]

        fig.add_trace(go.Scatter(
            x=[p["seasons"]    for p in group],
            y=[p["avg_vazeno"] for p in group],
            mode="markers+text",
            name=TEAM_NAMES.get(primary, primary),
            marker=dict(
                size=sizes, color=color, opacity=0.80,
                line=dict(width=1.2, color="white"),
            ),
            text=[p["name"].split()[-1] for p in group],
            textposition="top center",
            textfont=dict(size=9, color="#333333"),
            hovertext=hovers,
            hoverinfo="text",
        ))

    # Quadrant reference lines
    fig.add_vline(x=med_x, line_dash="dash", line_color="#CCCCCC",
                  line_width=1)
    fig.add_hline(y=med_y, line_dash="dash", line_color="#CCCCCC",
                  line_width=1)

    # Quadrant labels
    x_max = max(all_x) * 1.05
    y_max = max(all_y) * 1.05
    for (qx, qy, ha, va, label) in [
        (x_max, y_max,    "right", "top",    "Legendy"),
        (0,     y_max,    "left",  "top",    "Specialisté"),
        (0,     0,        "left",  "bottom", "Nováčci"),
        (x_max, 0,        "right", "bottom", "Stálice"),
    ]:
        fig.add_annotation(
            x=qx, y=qy, xref="x", yref="y",
            text=f"<i>{label}</i>",
            showarrow=False,
            font=dict(size=11, color="#BBBBBB"),
            xanchor=ha, yanchor=va,
        )

    fig.update_layout(
        title=(
            "Kariérní hodnota — zkušenosti vs. výkonnost<br>"
            "<sup>Velikost bodu = celkové Body Váženo</sup>"
        ),
        xaxis=dict(
            title="Odehráno sezón",
            gridcolor="#EEEEEE", zeroline=False,
        ),
        yaxis=dict(
            title="Ø Body Váženo / sezóna",
            gridcolor="#EEEEEE", zeroline=False,
        ),
        height=580,
        plot_bgcolor="#FAFAFA",
        paper_bgcolor="white",
        margin=dict(l=70, r=60, t=90, b=70),
        legend=dict(orientation="v", x=1.01, y=1, xanchor="left"),
    )
    return fig


# ── HTML assembly ────────────────────────────────────────────────────────────

_TAB_CSS = """
body { font-family: 'Segoe UI', 'DejaVu Sans', sans-serif;
       background:#f0f2f5; margin:0; padding:0; }
.header { background: linear-gradient(135deg, #1A6FBF 0%, #0D4A8A 100%);
          color:white; padding:20px 28px; }
.header h1 { margin:0 0 4px 0; font-size:1.6em; font-weight:700; }
.header p  { margin:0; opacity:0.82; font-size:0.95em; }
.tabs { display:flex; background:white;
        border-bottom:2px solid #dde1e7; padding:0 24px;
        box-shadow:0 2px 4px rgba(0,0,0,.06); }
.tab-btn { background:none; border:none; padding:13px 22px;
           cursor:pointer; font-size:14px; font-weight:500;
           color:#555; border-bottom:3px solid transparent;
           margin-bottom:-2px; transition:all .18s; }
.tab-btn:hover  { color:#1A6FBF; background:#f0f4ff; }
.tab-btn.active { color:#1A6FBF; border-bottom-color:#1A6FBF; }
.tab-content { display:none; padding:16px 8px; }
.tab-content.active { display:block; }
"""

_TAB_JS = """
function showTab(id, btn) {
  document.querySelectorAll('.tab-content').forEach(e=>e.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(e=>e.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  btn.classList.add('active');
}
"""

def generate_html(charts: list[tuple[str, str, str, go.Figure]],
                  output_path: str,
                  page_title: str = "TK Sport Kolovraty — Historická statistika",
                  header_title: str = "TK Sport Kolovraty",
                  header_subtitle: str = "Historická statistika dospělých — kariérní přehled"):
    """
    charts: list of (tab_id, emoji_icon, tab_title, figure)
    Writes a single self-contained HTML with plotly.js bundled.
    """
    parts = [f"""<!DOCTYPE html>
<html lang="cs">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{page_title}</title>
<style>{_TAB_CSS}</style>
</head>
<body>
<div class="header">
  <h1>{header_title}</h1>
  <p>{header_subtitle}</p>
</div>
<div class="tabs">
"""]
    for i, (tid, icon, title, _) in enumerate(charts):
        active = ' active' if i == 0 else ''
        parts.append(
            f'  <button class="tab-btn{active}" '
            f'onclick="showTab(\'{tid}\',this)">{icon} {title}</button>\n'
        )
    parts.append("</div>\n")

    for i, (tid, icon, title, fig) in enumerate(charts):
        active = ' active' if i == 0 else ''
        # Embed plotly.js only once (first chart); reuse for subsequent
        fig_html = pio.to_html(
            fig,
            include_plotlyjs=(i == 0),   # True = bundle full JS; False = reuse
            full_html=False,
            config={"responsive": True, "displayModeBar": True,
                    "modeBarButtonsToRemove": ["lasso2d","select2d"]},
            div_id=f"plot_{tid}",
        )
        parts.append(f'<div id="{tid}" class="tab-content{active}">\n{fig_html}\n</div>\n')

    parts.append(f"<script>{_TAB_JS}</script>\n</body>\n</html>")

    os.makedirs(Path(output_path).parent, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("".join(parts))


# ── Entry point ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Generate interactive career history charts.")
    parser.add_argument("--store", default=HISTORY_STORE,
                        help=f"Path to history_store.json (default: {HISTORY_STORE})")
    parser.add_argument("--out", default=OUTPUT,
                        help=f"Output HTML path (default: {OUTPUT})")
    args = parser.parse_args()

    if not Path(args.store).exists():
        print(f"ERROR: {args.store} not found. Run main.py first.")
        return

    print(f"Reading {args.store} ...")
    players = load_players(args.store)
    n = len(players)
    n_scraped = sum(1 for p in players if p["_seasons"])
    print(f"  {n} players  |  {n_scraped} with scraped year data")
    print()

    charts = [
        ("leaderboard", "🏆", "Pořadí",    fig_leaderboard(players)),
        ("timeline",    "📅", "Kariéra",    fig_timeline(players)),
        ("radar",       "🎯", "Srovnání",   fig_radar(players)),
        ("value",       "💡", "Hodnota",    fig_career_value(players)),
    ]

    for _, icon, title, _ in charts:
        print(f"  {icon} {title} — OK")

    print(f"\nWriting {args.out} ...")
    generate_html(charts, args.out)
    size_kb = Path(args.out).stat().st_size // 1024
    print(f"Done — {args.out}  ({size_kb} KB)")
    print("Open in any browser.")


if __name__ == "__main__":
    main()
