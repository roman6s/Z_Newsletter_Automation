#!/usr/bin/env python3
"""
IBM Z Newsletter Automation – Kommandozeilen-Interface
------------------------------------------------------
Fetches blog articles and events from the IBM Z DACH Community for a given
date range, summarizes articles in German using the Groq API, and generates
a filled newsletter PPTX based on the existing template.

Usage:
    python main.py

You will be prompted for:
  - Start date (YYYY-MM-DD)
  - End date (YYYY-MM-DD)
  - Issue number (e.g. 14)

Requirements:
    pip install -r requirements.txt

Groq API Key (free):
    https://console.groq.com
    Set as environment variable: export GROQ_API_KEY='your-key'
    Or edit config.py directly.
"""

import os
import sys
from datetime import date, datetime

from config import GROQ_API_KEY, OUTPUT_DIR, TEMPLATE_FILE


def prompt_date(label: str) -> date:
    while True:
        raw = input(f"{label} (YYYY-MM-DD): ").strip()
        try:
            return datetime.strptime(raw, "%Y-%m-%d").date()
        except ValueError:
            print("  Ungültiges Format. Bitte YYYY-MM-DD verwenden (z.B. 2026-01-01).")


def prompt_issue_number() -> str:
    raw = input("Issue-Nummer (z.B. 14): ").strip()
    return raw if raw else "?"


def main():
    print("=" * 55)
    print("   IBM Z Newsletter Automation")
    print("=" * 55)
    print()

    if not GROQ_API_KEY:
        print("FEHLER: Kein Groq API Key gefunden!")
        print("  Bitte setze ihn in config.py: GROQ_API_KEY = 'dein-key'")
        print("  Kostenlosen Key: https://console.groq.com")
        sys.exit(1)

    if not os.path.exists(TEMPLATE_FILE):
        print(f"FEHLER: Template-Datei nicht gefunden: {TEMPLATE_FILE}")
        sys.exit(1)

    print("Zeitraum für den Newsletter:")
    start_date = prompt_date("  Von")
    end_date   = prompt_date("  Bis")

    if start_date > end_date:
        print("FEHLER: Startdatum muss vor dem Enddatum liegen.")
        sys.exit(1)

    print()
    issue_number = prompt_issue_number()

    newsletter_month = end_date.month
    newsletter_year  = end_date.year

    print()
    print(f"Erstelle Newsletter für {start_date} bis {end_date} (Issue {issue_number})...")
    print()

    # Schritt 1: Artikel scrapen
    print("[1/4] Artikel von der IBM Community laden...")
    from scraper import fetch_articles_in_range
    articles = fetch_articles_in_range(start_date, end_date, verbose=True)

    if not articles:
        print()
        print("Keine Artikel für diesen Zeitraum gefunden.")
        print("Bitte überprüfe das Datum oder wähle einen anderen Zeitraum.")
        sys.exit(0)

    print(f"  → {len(articles)} Artikel gefunden.")
    print()

    # Schritt 2: Events scrapen (Upcoming: end_date + 1 Monat, max. 10)
    import calendar as _cal
    _m = end_date.month - 1 + 1
    events_start = end_date
    events_end = end_date.replace(
        year=end_date.year + _m // 12,
        month=_m % 12 + 1,
        day=min(end_date.day, _cal.monthrange(end_date.year + _m // 12, _m % 12 + 1)[1]),
    )
    print(f"[2/4] Upcoming Events laden ({events_start} bis {events_end})...")
    from event_scraper import fetch_events_in_range
    events = fetch_events_in_range(events_start, events_end, verbose=True)
    events_truncated = len(events) > 10
    if events_truncated:
        events = events[:10]
        print(f"  → Mehr als 10 Events gefunden – auf 10 begrenzt.")
    event_dicts = [
        {
            "title": ev.title,
            "event_date": ev.event_date,
            "time_str": ev.time_str,
            "location": ev.location,
            "description": ev.description,
            "url": ev.url,
        }
        for ev in events
    ]
    print(f"  → {len(event_dicts)} Events gefunden.")
    print()

    # Schritt 3: Artikel zusammenfassen
    print("[3/4] Artikel zusammenfassen (Groq API)...")
    from summarizer import summarize_articles
    summarized = summarize_articles(articles, verbose=True)
    print(f"  → {len(summarized)} Zusammenfassungen erstellt.")
    print()

    # Schritt 4: PPTX erstellen
    print("[4/4] Newsletter-PPTX erstellen...")
    from pptx_builder import build_newsletter
    output_path = build_newsletter(
        articles=summarized,
        month=newsletter_month,
        year=newsletter_year,
        issue_number=issue_number,
        events=event_dicts,
        events_truncated=events_truncated,
    )
    print(f"  → Gespeichert unter: {output_path}")
    print()
    print("=" * 55)
    print(f"  Fertig! Newsletter wurde erstellt:")
    print(f"  {os.path.abspath(output_path)}")
    print("=" * 55)


if __name__ == "__main__":
    main()
