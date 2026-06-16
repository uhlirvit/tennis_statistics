"""
Parser functions for cesky-tenis.cz team competition pages.

Two page types are parsed:
1. Competition/team page  (/soutez/{comp_id}?druzstvoCislo={team_no})
   -> list of rounds (date, opponent, our home/away, link to official record)
   -> season roster (name + player ID)
2. Official match record (/oficialni-zapis/{comp_id}?zapas={zapas_id})
   -> per-position results (singles/doubles, players, sets, win/loss)
"""

import re
from bs4 import BeautifulSoup


def extract_player_id(href: str) -> str | None:
    """Pull the numeric player ID out of an /hrac/{id}... href."""
    if not href:
        return None
    m = re.search(r"/hrac/(\d+)", href)
    return m.group(1) if m else None


def parse_roster(html: str) -> list[dict]:
    """Parse the 'Soupiska' (roster) table on a competition/team page."""
    soup = BeautifulSoup(html, "html.parser")
    roster = []
    for h2 in soup.find_all("h2", class_="in-box"):
        if "Soupiska" not in h2.get_text():
            continue
        box = h2.find_parent("div", class_="box")
        table = box.find("table", class_="table__table")
        for tr in table.find("tbody").find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 5:
                continue
            a = tds[1].find("a")
            if not a:
                continue
            roster.append({
                "roster_no": tds[0].get_text(strip=True),
                "name": a.get_text(strip=True),
                "player_id": extract_player_id(a.get("href", "")),
                "birth_year": tds[2].get_text(strip=True),
                "cz": tds[3].get_text(strip=True),
                "bh": tds[4].get_text(strip=True),
            })
    return roster


def parse_rounds(html: str, club_name: str) -> list[dict]:
    """
    Parse the 'Časový sled utkání' table -> one entry per round with
    date, opponent name, whether our club was home/away, and the zapas id
    needed to fetch the official record page.
    """
    soup = BeautifulSoup(html, "html.parser")
    rounds = []
    for h2 in soup.find_all("h2", class_="in-box"):
        if "sled" not in h2.get_text():
            continue
        box = h2.find_parent("div", class_="box")
        table = box.find("table", class_="table__table")
        for tr in table.find("tbody").find_all("tr"):
            tds = tr.find_all("td")
            if len(tds) < 9:
                continue
            kolo = tds[0].get_text(strip=True).rstrip(".")
            date = tds[1].get_text(strip=True)
            home_a = tds[3].find("a")
            away_a = tds[4].find("a")
            home_team = home_a.get_text(strip=True) if home_a else None
            away_team = away_a.get_text(strip=True) if away_a else None
            score = tds[5].get_text(strip=True)

            zapis_a = tds[8].find("a")
            zapis_href = zapis_a.get("href") if zapis_a else None
            zapas_id = None
            if zapis_href:
                m = re.search(r"zapas=(\d+)", zapis_href)
                zapas_id = m.group(1) if m else None

            if home_team == club_name:
                our_side = "D"
            elif away_team == club_name:
                our_side = "H"
            else:
                our_side = None  # shouldn't happen on our team's page

            rounds.append({
                "kolo": kolo,
                "date": date,
                "home_team": home_team,
                "away_team": away_team,
                "score": score,
                "our_side": our_side,
                "opponent": away_team if our_side == "D" else home_team,
                "zapas_id": zapas_id,
            })
    return rounds


def _cell_players(td) -> list[dict]:
    """Extract one or two (name, id) pairs from a Jmeno table cell."""
    players = []
    for a in td.find_all("a"):
        players.append({
            "name": a.get_text(strip=True),
            "player_id": extract_player_id(a.get("href", "")),
        })
    return players


def parse_match_record(html: str, club_name: str) -> dict:
    """
    Parse an /oficialni-zapis page -> match metadata + list of 9 (or fewer)
    position results, each tagged with which side is "our" club and the
    win/loss outcome for our side at that position.
    """
    soup = BeautifulSoup(html, "html.parser")

    # Home / away team names from the info box
    info_ps = soup.select(".box--white-10 .info")
    home_team = away_team = None
    for p in info_ps:
        span = p.find("span")
        a = p.find("a")
        if not span or not a:
            continue
        label = span.get_text(strip=True)
        if label == "Domácí":
            home_team = a.get_text(strip=True)
        elif label == "Hosté":
            away_team = a.get_text(strip=True)

    # Date / venue
    date = venue = None
    for p in soup.select(".row .info"):
        span = p.find("span")
        if not span:
            continue
        label = span.get_text(strip=True)
        if label == "Datum":
            date = p.get_text(strip=True).replace("Datum", "").strip()
        elif label == "Místo konání":
            venue = p.get_text(strip=True).replace("Místo konání", "").strip()

    if home_team == club_name:
        our_side = "D"
    elif away_team == club_name:
        our_side = "H"
    else:
        our_side = None

    table = soup.find("table", class_="table__table")
    positions = []
    for tr in table.find("tbody").find_all("tr"):
        if "winner" in tr.get("class", []):
            continue
        tds = tr.find_all("td")
        if len(tds) < 14:
            continue

        pos_num = tds[0].get_text(strip=True).rstrip(".")
        home_players = _cell_players(tds[1])
        away_players = _cell_players(tds[3])

        body_d = tds[12].get_text(strip=True)
        body_h = tds[13].get_text(strip=True)

        match_type = "debl" if len(home_players) == 2 else "singl"

        if our_side == "D":
            our_players, opp_players = home_players, away_players
            our_body, opp_body = body_d, body_h
        elif our_side == "H":
            our_players, opp_players = away_players, home_players
            our_body, opp_body = body_h, body_d
        else:
            our_players, opp_players, our_body, opp_body = [], [], None, None

        if our_body == "1":
            result = "W"
        elif our_body == "0":
            result = "L"
        else:
            result = None  # unexpected / not played

        positions.append({
            "position": pos_num,
            "type": match_type,
            "our_players": our_players,
            "opponent_players": opp_players,
            "result": result,
            "sets": [tds[5].get_text(strip=True), tds[6].get_text(strip=True), tds[7].get_text(strip=True)],
            "games": (tds[8].get_text(strip=True), tds[9].get_text(strip=True)),
        })

    return {
        "date": date,
        "venue": venue,
        "home_team": home_team,
        "away_team": away_team,
        "our_side": our_side,
        "positions": positions,
    }
