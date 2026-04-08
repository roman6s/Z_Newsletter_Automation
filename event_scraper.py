"""
Scraper for IBM Z DACH Community Events Calendar.

Fetches events from the IBM Community Higher Logic platform.
Returns events filtered to a given date range.
"""

import re
import time
from dataclasses import dataclass
from datetime import date, datetime
from typing import List, Optional, Tuple

import requests
from bs4 import BeautifulSoup

from config import EVENTS_URL

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "de-DE,de;q=0.9,en;q=0.8",
}

# Month name mapping for date parsing (EN + DE)
_MONTH_MAP = {
    "january": 1, "february": 2, "march": 3, "april": 4,
    "may": 5, "june": 6, "july": 7, "august": 8,
    "september": 9, "october": 10, "november": 11, "december": 12,
    "januar": 1, "februar": 2, "märz": 3,
    "mai": 5, "juni": 6, "juli": 7,
    "oktober": 10, "dezember": 12,
}


@dataclass
class Event:
    title: str
    event_date: date
    time_str: str = ""
    location: str = ""
    description: str = ""
    url: str = ""


# ── Date parsing ──────────────────────────────────────────────────────────────

def _parse_date(text: str) -> Optional[date]:
    """Try to extract a date from a human-readable string."""
    text = text.strip()
    # Remove leading weekday names
    text = re.sub(
        r'^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|'
        r'Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag),?\s+',
        '', text, flags=re.I,
    )

    # Standard format attempts
    for fmt in (
        "%B %d, %Y",   # April 15, 2026
        "%b %d, %Y",   # Apr 15, 2026
        "%d. %B %Y",   # 15. April 2026
        "%d.%m.%Y",    # 15.04.2026
        "%Y-%m-%d",    # 2026-04-15
        "%m/%d/%Y",    # 04/15/2026
        "%d/%m/%Y",    # 15/04/2026
    ):
        try:
            return datetime.strptime(text[: len(fmt) + 6], fmt).date()
        except (ValueError, IndexError):
            pass

    # Regex fallback: "Month Day, Year" or "Day. Month Year"
    m = re.search(r'(\d{1,2})\.\s+(\w+)\s+(\d{4})', text)
    if m:
        month_num = _MONTH_MAP.get(m.group(2).lower())
        if month_num:
            try:
                return date(int(m.group(3)), month_num, int(m.group(1)))
            except ValueError:
                pass

    m = re.search(r'(\w+)\s+(\d{1,2}),?\s+(\d{4})', text)
    if m:
        month_num = _MONTH_MAP.get(m.group(1).lower())
        if month_num:
            try:
                return date(int(m.group(3)), month_num, int(m.group(2)))
            except ValueError:
                pass

    return None


# ── HTML field extraction ─────────────────────────────────────────────────────

def _get_hidden_fields(soup: BeautifulSoup) -> dict:
    """Collect ALL hidden input fields – required for ASP.NET postback (incl. CSRF token)."""
    fields = {}
    for inp in soup.find_all("input", type="hidden"):
        name = inp.get("name", "")
        if name:
            fields[name] = inp.get("value", "")
    return fields


def _find_next_pager_target(
    soup: BeautifulSoup,
    current_page: int,
    visited: set,
) -> Tuple[Optional[str], Optional[str]]:
    """Return (target, arg) for the next page, or (None, None) if exhausted."""
    next_page_text = str(current_page + 1)
    arrow_texts = {">", ">>", "»", "›", "Next", "Nächste"}
    arrow_candidate: Tuple[Optional[str], Optional[str]] = (None, None)

    for a in soup.find_all("a", href=True):
        href = a.get("href", "")
        m = re.search(r"__doPostBack\('([^']+)','([^']*)'\)", href)
        if not m:
            continue
        target, arg = m.group(1), m.group(2)
        if target in visited:
            continue
        text = a.get_text(strip=True)
        if text == next_page_text:
            return target, arg
        if text in arrow_texts and arrow_candidate == (None, None):
            arrow_candidate = (target, arg)

    return arrow_candidate


# ── Event extraction strategies ───────────────────────────────────────────────

def _extract_events_strategy_aspnet(soup: BeautifulSoup) -> List[dict]:
    """
    IBM Community Higher Logic events page:
      - Title link: id="MainCopy_ctl06_lvSearchResults_hypTitle_N"
      - Date/time:  id="MainCopy_ctl06_lvSearchResults_pnlCalendarLocation_N"
                    text format: "Tue April 07, 2026|10:00 AM - 10:30 AM ET"
      - Description:id="MainCopy_ctl06_lvSearchResults_pDescription_N"
    """
    events = []
    title_pattern = re.compile(r"MainCopy_ctl06_lvSearchResults_hypTitle_(\d+)$", re.I)

    for a_tag in soup.find_all("a", id=title_pattern):
        m = title_pattern.search(a_tag["id"])
        if not m:
            continue
        idx = m.group(1)
        title = a_tag.get("title", a_tag.get_text(strip=True))
        url = a_tag.get("href", "")
        if not title:
            continue

        event_date = None
        time_str = ""
        description = ""

        # Date / time / location: "Tue April 07, 2026|10:00 AM - 10:30 AM ET"
        loc_el = soup.find(id=f"MainCopy_ctl06_lvSearchResults_pnlCalendarLocation_{idx}")
        if loc_el:
            raw = loc_el.get_text(strip=True)
            if "|" in raw:
                date_part, time_part = raw.split("|", 1)
                event_date = _parse_date(date_part.strip())
                time_str = time_part.strip()
            else:
                event_date = _parse_date(raw)

        # Description
        desc_el = soup.find(id=f"MainCopy_ctl06_lvSearchResults_pDescription_{idx}")
        if desc_el:
            description = desc_el.get_text(strip=True)[:300]

        events.append({
            "title": title,
            "url": url,
            "event_date": event_date,
            "time_str": time_str,
            "location": "",
            "description": description,
        })

    return events


def _extract_events_from_page(soup: BeautifulSoup) -> List[dict]:
    """Extract all events from the current page."""
    return _extract_events_strategy_aspnet(soup)


# ── Public API ────────────────────────────────────────────────────────────────

def fetch_events_in_range(
    start_date: date,
    end_date: date,
    verbose: bool = True,
) -> List[Event]:
    """
    Fetch IBM Z DACH Community events within [start_date, end_date].
    Returns a sorted list of Event objects.
    Never raises – returns [] on any error.
    """
    session = requests.Session()
    session.headers.update(HEADERS)

    if verbose:
        print("  Events-Seite laden...")

    try:
        resp = session.get(EVENTS_URL, timeout=30)
        resp.raise_for_status()
    except Exception as exc:
        if verbose:
            print(f"  Warnung: Events-Seite nicht erreichbar: {exc}")
        return []

    soup = BeautifulSoup(resp.text, "html.parser")

    raw: List[dict] = []
    visited: set = set()
    current_page = 1
    max_pages = 30

    raw.extend(_extract_events_from_page(soup))
    if verbose:
        print(f"  Seite 1: {len(raw)} Einträge gefunden")

    while current_page < max_pages:
        target, arg = _find_next_pager_target(soup, current_page, visited)
        if target is None:
            break

        visited.add(target)
        hidden = _get_hidden_fields(soup)
        hidden["__EVENTTARGET"] = target
        hidden["__EVENTARGUMENT"] = arg or ""

        try:
            resp = session.post(EVENTS_URL, data=hidden, timeout=30)
            resp.raise_for_status()
        except Exception as exc:
            if verbose:
                print(f"  Warnung: Seite {current_page + 1} nicht ladbar: {exc}")
            break

        soup = BeautifulSoup(resp.text, "html.parser")
        page_raw = _extract_events_from_page(soup)
        if not page_raw:
            break

        raw.extend(page_raw)
        current_page += 1
        if verbose:
            print(f"  Seite {current_page}: {len(page_raw)} Einträge gefunden")

        # Stop if all dated entries on this page are already past end_date
        dated = [e for e in page_raw if e.get("event_date")]
        if dated and all(e["event_date"] > end_date for e in dated):
            break

        time.sleep(0.5)

    # Deduplicate, filter, convert
    seen_keys: set = set()
    events: List[Event] = []

    for e in raw:
        key = e.get("url") or e.get("title", "")
        if not key or key in seen_keys:
            continue
        seen_keys.add(key)

        evt_date = e.get("event_date")
        if evt_date is None:
            continue
        if evt_date < start_date or evt_date > end_date:
            continue

        events.append(Event(
            title=e["title"],
            event_date=evt_date,
            time_str=e.get("time_str", ""),
            location=e.get("location", ""),
            description=e.get("description", ""),
            url=e.get("url", ""),
        ))

    events.sort(key=lambda ev: ev.event_date)

    if verbose:
        print(f"  → {len(events)} Events im Zeitraum gefunden.")

    return events
