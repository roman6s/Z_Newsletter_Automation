"""
Scraper for IBM Z DACH Community Events.

Uses the Higher Logic JSON API that powers the community events page.
No pagination needed – the API returns all upcoming events in one call.
"""

import re
import time
from dataclasses import dataclass
from datetime import date, datetime
from typing import List, Optional

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

# Content key for the IBM Z DACH community events widget
CONTENT_KEY = "4fb54aeb-7d2d-40b3-a33e-019746231df1"
API_PATH = "/higherlogic/ocapi/oc/events/communityEventsList/getEventsList"

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


def _parse_date(text: str) -> Optional[date]:
    """Parse a human-readable date string like 'Thursday, January 1, 2026'."""
    text = text.strip()
    # Strip leading weekday
    text = re.sub(
        r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+",
        "", text, flags=re.I,
    )

    for fmt in ("%B %d, %Y", "%b %d, %Y", "%d. %B %Y", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text[: len(fmt) + 6], fmt).date()
        except (ValueError, IndexError):
            pass

    m = re.search(r"(\w+)\s+(\d{1,2}),?\s+(\d{4})", text)
    if m:
        month_num = _MONTH_MAP.get(m.group(1).lower())
        if month_num:
            try:
                return date(int(m.group(3)), month_num, int(m.group(2)))
            except ValueError:
                pass

    return None


def _clean_text(text: str) -> str:
    """Collapse all whitespace into single spaces and strip."""
    text = re.sub(r"[ \t]+", " ", text)        # multiple spaces/tabs → single space
    text = re.sub(r"\n{3,}", "\n\n", text)      # 3+ newlines → double newline
    text = re.sub(r" ?\n ?", "\n", text)        # spaces around newlines
    return text.strip()


def _get_form_token(session: requests.Session) -> str:
    """Fetch the listing page and extract the CSRF form token."""
    resp = session.get(EVENTS_URL, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    inp = soup.find("input", {"name": "__HL-RequestVerificationToken"})
    return inp["value"] if inp else ""


def fetch_events_in_range(
    start_date: date,
    end_date: date,
    verbose: bool = True,
) -> List[Event]:
    """
    Fetch IBM Z DACH Community events within [start_date, end_date].
    Returns a sorted list of Event objects (max 10).
    Never raises – returns [] on any error.
    """
    session = requests.Session()
    session.headers.update(HEADERS)

    if verbose:
        print("  Events laden (API)...")

    try:
        form_token = _get_form_token(session)
    except Exception as exc:
        if verbose:
            print(f"  Warnung: Events-Seite nicht erreichbar: {exc}")
        return []

    base = EVENTS_URL.split("/community/")[0]  # e.g. https://community.ibm.com
    api_url = base + API_PATH

    payload = {
        "contentKey": CONTENT_KEY,
        "maxEvents": 50,
        "showPastEvents": False,
        "showFutureEvents": True,
        "showFeaturedImage": True,
        "showEventDescription": True,
        "showEventLocation": True,
    }

    try:
        api_resp = session.post(
            api_url,
            json=payload,
            timeout=30,
            headers={
                "Content-Type": "application/json",
                "Referer": EVENTS_URL,
                "RequestVerificationFormToken": form_token,
            },
        )
        api_resp.raise_for_status()
        data = api_resp.json()
    except Exception as exc:
        if verbose:
            print(f"  Warnung: Events-API nicht erreichbar: {exc}")
        return []

    raw_events = data.get("data", {}).get("futureCommunityEventsList", [])

    if verbose:
        print(f"  {len(raw_events)} Events von API erhalten")

    events: List[Event] = []
    for ev in raw_events:
        evt_date = _parse_date(ev.get("startDate", ""))
        if evt_date is None:
            continue
        if evt_date < start_date or evt_date > end_date:
            continue

        location = ev.get("location", "") or ""
        if location.lower() in ("not specified", "nicht angegeben", ""):
            location = "Online" if ev.get("isOnlineEvent") or ev.get("isConferenceCallEvent") else ""

        raw_desc = ev.get("description", "") or ""
        description = _clean_text(raw_desc)
        # Limit to ~300 chars, cut at sentence boundary
        if len(description) > 300:
            cut = description[:300]
            last_period = max(cut.rfind(". "), cut.rfind(".\n"))
            description = cut[: last_period + 1] if last_period > 100 else cut.rstrip() + "…"

        events.append(Event(
            title=ev.get("title", "").strip(),
            event_date=evt_date,
            time_str=ev.get("dateRange", "").strip(),
            location=location.strip(),
            description=description,
            url=ev.get("linkToView", "").strip(),
        ))

    events.sort(key=lambda e: e.event_date)

    if verbose:
        print(f"  → {len(events)} Events im Zeitraum [{start_date} – {end_date}]")

    return events
