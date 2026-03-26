#!/usr/bin/env python3
"""
IBM Z Newsletter Automation
----------------------------
Fetches blog articles from the IBM Z DACH Community for a given date range,
summarizes them in German using the Gemini API, and generates a filled
newsletter PPTX based on the existing template.

Usage:
    python main.py

You will be prompted for:
  - Start date (YYYY-MM-DD)
  - End date (YYYY-MM-DD)
  - Issue number (e.g. 14)

Requirements:
  pip install requests beautifulsoup4 google-generativeai python-pptx

Gemini API Key (free):
  https://aistudio.google.com/apikey
  Set as environment variable: export GEMINI_API_KEY='your-key'
  Or edit config.py directly.
"""

import os
import sys
from datetime import date, datetime

from config import GROQ_API_KEY, OUTPUT_DIR


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

    # Check API key
    if not GROQ_API_KEY:
        print("FEHLER: Kein Groq API Key gefunden!")
        print("  Bitte setze ihn in config.py: GROQ_API_KEY = 'dein-key'")
        print("  Kostenlosen Key: https://console.groq.com")
        sys.exit(1)

    # Check template
    if not os.path.exists("Oktober Newsletter.pptx"):
        print("FEHLER: Template-Datei 'Oktober Newsletter.pptx' nicht gefunden!")
        print("  Bitte stelle sicher, dass du das Programm im Ordner")
        print("  IBM_Z_Newsletter_Automation ausführst.")
        sys.exit(1)

    # Get date range
    print("Zeitraum für den Newsletter:")
    start_date = prompt_date("  Von")
    end_date = prompt_date("  Bis")

    if start_date > end_date:
        print("FEHLER: Startdatum muss vor dem Enddatum liegen.")
        sys.exit(1)

    # Get issue number
    print()
    issue_number = prompt_issue_number()

    # Determine month/year for the newsletter header
    # Use the end date's month/year
    newsletter_month = end_date.month
    newsletter_year = end_date.year

    print()
    print(f"Erstelle Newsletter für {start_date} bis {end_date} (Issue {issue_number})...")
    print()

    # Step 1: Scrape articles
    print("[1/3] Artikel von der IBM Community laden...")
    from scraper import fetch_articles_in_range
    articles = fetch_articles_in_range(start_date, end_date, verbose=True)

    if not articles:
        print()
        print("Keine Artikel für diesen Zeitraum gefunden.")
        print("Bitte überprüfe das Datum oder wähle einen anderen Zeitraum.")
        sys.exit(0)

    print(f"  → {len(articles)} Artikel gefunden.")
    print()

    # Step 2: Summarize
    print("[2/3] Artikel zusammenfassen (Gemini API)...")
    from summarizer import summarize_articles
    summarized = summarize_articles(articles, verbose=True)
    print(f"  → {len(summarized)} Zusammenfassungen erstellt.")
    print()

    # Step 3: Build PPTX
    print("[3/3] Newsletter-PPTX erstellen...")
    from pptx_builder import build_newsletter
    output_path = build_newsletter(
        articles=summarized,
        month=newsletter_month,
        year=newsletter_year,
        issue_number=issue_number,
    )
    print(f"  → Gespeichert unter: {output_path}")
    print()
    print("=" * 55)
    print(f"  Fertig! Newsletter wurde erstellt:")
    print(f"  {os.path.abspath(output_path)}")
    print("=" * 55)


if __name__ == "__main__":
    main()
