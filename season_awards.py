"""
season_awards.py — Interactive per-team season leaderboard (Pořadí / Ocenění).

Generates a single HTML file with one TAB PER TEAM, across ALL categories
in a season (dospělí, dorost, mladší žáci, baby, ...). Inside each tab,
a dropdown switches between the same six award categories used in the
season workbook's Ocenění sheet:

    MVP · Singlový specialista · Deblový specialista ·
    Nejvíce odehráno · Singl výhry · Debl výhry

Colour scheme:
    - Dospělí teams keep their usual per-team colours (A=blue, B=orange,
      C=green, D=purple).
    - Each youth category gets its OWN distinct colour family (dorost=teal,
      mladší žáci=raspberry, baby=gold, ...), so tabs are visually grouped
      by category at a glance.

Reads directly from xlsx files, so it reflects any manual edits.

Usage:
    # Single category (unchanged from before)
    python3 season_awards.py output/2026_dospeli.xlsx

    # Explicit multiple files — combined into one HTML
    python3 season_awards.py output/2026_dospeli.xlsx output/2026_dorost.xlsx output/2026_baby.xlsx

    # Season shorthand — auto-discovers output/2026_*.xlsx
    python3 season_awards.py 2026

    python3 season_awards.py 2026 --out output/season_2026_all.html

Requires: pip install plotly openpyxl
"""

import argparse
import colorsys
from pathlib import Path

import plotly.graph_objects as go

from visualize import extract_workbook, _team_color, C_BLUE
from history_charts import generate_html

# ── Award categories — same labels/order as the season workbook's Ocenění sheet
CATEGORIES = [
    ("vazeno",   "MVP",                    False),
    ("us_singl", "Singlový specialista",   True),
    ("us_debl",  "Deblový specialista",    True),
    ("celkem",   "Nejvíce odehráno",       False),
    ("vs",       "Singl výhry",            False),
    ("vd",       "Debl výhry",             False),
]

# ── Category metadata ───────────────────────────────────────────────────────
CATEGORY_ORDER = ["dospeli", "dorost", "mladsi", "zaci", "starsi", "mladez", "baby"]

CATEGORY_DISPLAY = {
    "dospeli": "Dospělí",
    "dorost":  "Dorost",
    "mladsi":  "Mladší žáci",
    "zaci":    "Žáci",
    "starsi":  "Starší žáci",
    "mladez":  "Mládež",
    "baby":    "Babytenis",
}

CATEGORY_ICONS = {
    "dospeli": "🎾",
    "dorost":  "🌱",
    "mladsi":  "🧒",
    "zaci":    "🧒",
    "starsi":  "🧑",
    "mladez":  "👦",
    "baby":    "🍼",
}

# Base hue per youth category — chosen to NOT clash with the dospělí
# per-team palette (blue/orange/green/purple from visualize.TEAM_COLORS)
CATEGORY_COLORS = {
    "dorost":  "#16A085",   # teal
    "mladsi":  "#C2185B",   # raspberry
    "zaci":    "#C2185B",   # same family as mladsi
    "starsi":  "#5D4037",   # brown
    "mladez":  "#34495E",   # slate
    "baby":    "#D4AC0D",   # gold
}
DEFAULT_CATEGORY_COLOR = "#888888"

# Shade multipliers applied (via HSL lightness) for multiple teams within
# the SAME youth category, e.g. if "mladsi" ever splits into A/B
SHADE_FACTORS = [1.0, 0.65, 1.35, 0.5]


def _hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb: tuple[float, float, float]) -> str:
    return "#%02X%02X%02X" % tuple(max(0, min(255, int(round(c)))) for c in rgb)


def _shade(hex_color: str, factor: float) -> str:
    """Lighten (factor>1) or darken (factor<1) a hex colour via HSL lightness."""
    r, g, b = _hex_to_rgb(hex_color)
    h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    l = max(0.18, min(0.82, l * factor))
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return _rgb_to_hex((r2 * 255, g2 * 255, b2 * 255))


def _resolve_color(category_key: str, label: str, team_index: int) -> str:
    """
    Pick a bar colour for a tab.
      - dospeli: keep the familiar per-team palette (A/B/C/D)
      - other categories: a category-specific hue, shaded per team if
        more than one team exists within that category
    """
    if category_key == "dospeli":
        return _team_color(label) if label not in ("", "All Club") else C_BLUE

    base = CATEGORY_COLORS.get(category_key, DEFAULT_CATEGORY_COLOR)
    if label in ("", "All Club"):
        return base
    return _shade(base, SHADE_FACTORS[team_index % len(SHADE_FACTORS)])


# ── Leaderboard figure (same as before — category-agnostic) ────────────────

def fig_team_leaderboard(summary: list[dict], tab_title: str, color: str) -> go.Figure:
    """
    Horizontal bar chart for one team, with a dropdown to switch between
    the 6 award categories. Only players who actually played are shown.
    """
    traces  = []
    buttons = []

    for i, (field, label, is_pct) in enumerate(CATEGORIES):
        valid = sorted(
            [p for p in summary if p.get("celkem", 0) > 0],
            key=lambda p: (p.get(field) or 0),
        )
        names  = [p["name"] for p in valid]
        values = [(p.get(field) or 0) * (100 if is_pct else 1) for p in valid]
        texts = [
            f"{v:.0f}%" if is_pct else
            (f"{v:.1f}" if isinstance(v, float) and not float(v).is_integer() else f"{int(v)}")
            for v in values
        ]
        hovers = [
            f"<b>{p['name']}</b><br>"
            f"Odehráno: {p['celkem']}  (singl {p['os']} / debl {p['od']})<br>"
            f"Výhry: singl {p['vs']} · debl {p['vd']}<br>"
            f"Body Suma: {p['suma']}  |  Body Váženo: {p['vazeno']:.1f}"
            for p in valid
        ]

        traces.append(go.Bar(
            x=values, y=names,
            orientation="h",
            marker_color=color, opacity=0.85,
            text=texts, textposition="outside",
            textfont=dict(size=10),
            hovertext=hovers, hoverinfo="text",
            visible=(i == 0),
            name=label,
            showlegend=False,
        ))

        x_title = f"{label} (%)" if is_pct else label
        buttons.append(dict(
            label=label,
            method="update",
            args=[
                {"visible": [j == i for j in range(len(CATEGORIES))]},
                {"xaxis.title": x_title,
                 "title.text":  f"Pořadí — {label}  ·  {tab_title}"},
            ],
        ))

    n_players = len([p for p in summary if p.get("celkem", 0) > 0]) or 1
    fig = go.Figure(data=traces)
    fig.update_layout(
        title=f"Pořadí — MVP  ·  {tab_title}",
        height=max(420, n_players * 28 + 130),
        xaxis_title="MVP",
        yaxis=dict(tickfont=dict(size=13, color="#1A1A1A"), automargin=True),
        plot_bgcolor="#FAFAFA",
        paper_bgcolor="white",
        margin=dict(l=20, r=40, t=70, b=50),
        updatemenus=[dict(
            type="dropdown",
            direction="down",
            buttons=buttons,
            x=0.0, xanchor="left",
            y=1.08, yanchor="top",
            bgcolor="white",
            bordercolor="#CCCCCC",
            font=dict(size=12),
            showactive=True,
        )],
    )
    return fig


# ── File / category parsing ─────────────────────────────────────────────────

def _parse_stem(path: str) -> tuple[str, str]:
    """'output/2026_dospeli.xlsx' → ('2026', 'dospeli')"""
    stem = Path(path).stem
    parts = stem.split("_", 1)
    season   = parts[0] if parts else stem
    category = parts[1] if len(parts) > 1 else ""
    return season, category


def _category_sort_key(path: str) -> int:
    _, category = _parse_stem(path)
    try:
        return CATEGORY_ORDER.index(category.lower())
    except ValueError:
        return len(CATEGORY_ORDER)


# ── Entry point ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Generate interactive season leaderboard (Pořadí), "
                    "across one or more categories.")
    parser.add_argument(
        "files", nargs="+",
        help="Season xlsx file(s), OR a single season year (e.g. 2026) "
             "to auto-discover all output/2026_*.xlsx files.")
    parser.add_argument("--output-dir", default="output",
                        help="Directory to search for the season-year shorthand "
                             "(default: output)")
    parser.add_argument("--out", default=None,
                        help="Output HTML path (default: output/<season>_awards.html)")
    args = parser.parse_args()

    # Season-year shorthand: a single numeric argument
    is_combined = False
    if len(args.files) == 1 and args.files[0].isdigit():
        season_arg = args.files[0]
        found = sorted(Path(args.output_dir).glob(f"{season_arg}_*.xlsx"))
        found = [f for f in found
                if "history" not in f.name and "test" not in f.name
                and "_awards" not in f.name]
        if not found:
            print(f"No files matching {season_arg}_*.xlsx found in {args.output_dir}/")
            return
        files = [str(f) for f in found]
        is_combined = True
        print(f"Auto-discovered {len(files)} file(s) for season {season_arg}:")
        for f in files:
            print(f"  {f}")
        print()
    else:
        files = args.files
        is_combined = len(files) > 1

    files = sorted(files, key=_category_sort_key)

    all_charts   = []
    seasons_seen = []
    cats_seen    = []

    for f in files:
        season, category = _parse_stem(f)
        if season not in seasons_seen:
            seasons_seen.append(season)
        cat_key     = category.lower()
        cat_display = CATEGORY_DISPLAY.get(cat_key, category.capitalize() if category else "Tým")
        if cat_display not in cats_seen:
            cats_seen.append(cat_display)

        print(f"Reading {f} ...")
        teams = extract_workbook(f)
        if not teams:
            print("  No recognisable sheets found — skipping.")
            continue

        team_labels = sorted(t for t in teams if t not in ("All Club", ""))
        if "" in teams:
            team_labels.append("")
        if "All Club" in teams:
            team_labels.append("All Club")

        lettered_idx = 0   # index among actual A/B/C/D-style teams, for shading
        for label in team_labels:
            data    = teams[label]
            summary = data.get("summary", [])
            if not summary:
                continue

            if label == "All Club":
                tab_title = f"{cat_display} Klub"
                icon = "🏆"
            elif label == "":
                tab_title = cat_display
                icon = CATEGORY_ICONS.get(cat_key, "🎾")
            else:
                tab_title = f"{cat_display} {label}"
                icon = "🏅"

            color = _resolve_color(cat_key, label, lettered_idx)
            if label not in ("", "All Club"):
                lettered_idx += 1

            fig = fig_team_leaderboard(summary, tab_title, color)
            slug = (label or "main").replace(" ", "")
            tab_id = f"{cat_key or category}_{slug}"
            all_charts.append((tab_id, icon, tab_title, fig))
            print(f"  {icon} {tab_title} — OK ({len(summary)} hráčů)")

    if not all_charts:
        print("\nNo teams with players found — nothing to write.")
        return

    season = seasons_seen[0] if seasons_seen else "season"
    if len(seasons_seen) > 1:
        print(f"\nWARNING: files span multiple seasons {seasons_seen} — "
              f"using '{season}' for the output filename.")

    out_path = args.out or (
        f"{args.output_dir}/{season}_awards.html" if is_combined
        else f"{args.output_dir}/{Path(files[0]).stem}_awards.html"
    )
    subtitle = f"Pořadí — sezóna {season} · " + ", ".join(cats_seen)

    print(f"\nWriting {out_path} ...")
    generate_html(
        all_charts, out_path,
        page_title=f"TK Sport Kolovraty — Pořadí {season}",
        header_title="TK Sport Kolovraty",
        header_subtitle=subtitle,
    )
    size_kb = Path(out_path).stat().st_size // 1024
    print(f"Done — {out_path}  ({size_kb} KB, {len(all_charts)} tabs)")
    print("Open in any browser.")


if __name__ == "__main__":
    main()
