"""
Summarizes blog articles in German using the Groq API (free tier).
Get your free API key at: https://console.groq.com
"""

import time

from groq import Groq

from config import GROQ_API_KEY, SUMMARY_LANGUAGE
from scraper import Article

MODEL = "llama-3.3-70b-versatile"
MAX_RETRIES = 3


def _build_client():
    if not GROQ_API_KEY:
        raise ValueError(
            "Kein Groq API Key gefunden!\n"
            "Kostenlosen Key bekommst du unter: https://console.groq.com\n"
            "Dann in config.py eintragen: GROQ_API_KEY = 'dein-key'"
        )
    return Groq(api_key=GROQ_API_KEY)


def summarize_article(article: Article, client=None) -> str:
    """Returns a German summary of the article (3-5 sentences)."""
    if client is None:
        client = _build_client()

    prompt = f"""Du erstellst Zusammenfassungen für einen IBM Z Newsletter auf {SUMMARY_LANGUAGE}.

Artikel-Titel: {article.title}
Autor: {article.author}

Artikel-Inhalt:
{article.full_text[:5000]}

Aufgabe: Schreibe eine strukturierte Zusammenfassung auf Deutsch mit 2-3 Absätzen.
- Trenne die Absätze jeweils mit einer Leerzeile (Absatz 1, Leerzeile, Absatz 2, ...)
- Erster Absatz: Worum geht es? (1-2 Sätze)
- Zweiter Absatz: Wichtigste Punkte oder Erkenntnisse (2-4 Sätze – gerne ausführlich wenn der Artikel viel hergibt)
- Optionaler dritter Absatz: Relevanz oder Ausblick (1-2 Sätze)
- Finde eine gute Balance: nicht zu knapp, aber auch nicht ausschweifend
- Professioneller, informativer Ton
- Falls der Text nicht auf Deutsch ist, übersetze die Kernaussagen
- Schreibe NUR die Zusammenfassung, ohne Einleitung oder Überschriften"""

    for attempt in range(MAX_RETRIES):
        try:
            response = client.chat.completions.create(
                model=MODEL,
                messages=[{"role": "user", "content": prompt}],
                max_tokens=400,
                temperature=0.3,
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            err = str(e)
            if "429" in err or "rate_limit" in err.lower():
                wait = 30 * (attempt + 1)
                print(f"    Rate limit, warte {wait}s...")
                time.sleep(wait)
            else:
                return f"[Zusammenfassung nicht verfügbar: {e}]"

    return "[Zusammenfassung nicht verfügbar: Rate limit nach mehreren Versuchen]"


def summarize_articles(articles: list, verbose: bool = True) -> list:
    """Returns list of dicts with keys: title, author, url, published, summary."""
    client = _build_client()
    results = []

    for i, article in enumerate(articles):
        if verbose:
            print(f"  Zusammenfassung {i+1}/{len(articles)}: {article.title[:60]}...")
        summary = summarize_article(article, client)
        results.append({
            "title": article.title,
            "author": article.author,
            "url": article.url,
            "published": article.published,
            "summary": summary,
            "image_url": article.image_url,
        })

    return results
