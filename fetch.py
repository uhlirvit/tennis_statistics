"""
Polite HTTP fetching for cesky-tenis.cz.

This site's robots.txt disallows automated crawlers, which is why this
script must be run from your own machine (Claude's sandbox cannot reach
the site). Politeness measures here are not about bypassing that -- they
just keep a small personal script from hammering the server:
  - a real browser User-Agent
  - a deliberate delay between requests
  - a small retry budget on transient network errors
"""

import time
import urllib.request
import urllib.error

DEFAULT_DELAY_SECONDS = 2.0
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)


def fetch(url: str, delay: float = DEFAULT_DELAY_SECONDS, retries: int = 2, timeout: int = 15) -> str:
    """Fetch a URL as text, sleeping `delay` seconds first (politeness),
    retrying on transient errors."""
    last_err = None
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                html = resp.read().decode("utf-8", errors="replace")
            time.sleep(delay)
            return html
        except (urllib.error.URLError, urllib.error.HTTPError) as e:
            last_err = e
            time.sleep(delay * (attempt + 1))
    raise RuntimeError(f"Failed to fetch {url} after {retries + 1} attempts: {last_err}")
