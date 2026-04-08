"""
IBM Z Newsletter Automation – Streamlit Frontend
Run with: streamlit run app.py
"""

import io
import json
import os
from datetime import date
from pathlib import Path

import streamlit as st

# ── Lokale Konfiguration speichern/laden ──────────────────────────────────────
CONFIG_FILE = Path(".saved_config.json")

def load_saved_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text())
        except Exception:
            pass
    return {}

def save_config(data: dict):
    try:
        CONFIG_FILE.write_text(json.dumps(data))
    except Exception:
        pass

saved = load_saved_config()

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IBM Z Newsletter",
    page_icon="📰",
    layout="centered",
)

# ── Header ────────────────────────────────────────────────────────────────────
st.title("📰 IBM Z Newsletter Automation")
st.caption("Erstellt automatisch einen Newsletter aus den aktuellen IBM Z DACH Blog-Artikeln.")
st.divider()

# ── Sidebar: API-Einstellungen ─────────────────────────────────────────────────
with st.sidebar:
    st.header("🔑 API-Einstellungen")
    st.markdown("Einmalig ausfüllen – wird lokal gespeichert.")

    groq_key = st.text_input(
        "Groq API Key",
        value=saved.get("groq_key", ""),
        type="password",
        placeholder="gsk_...",
        help="Kostenlosen Key erstellen unter console.groq.com",
    )

    model = st.selectbox(
        "Modell",
        options=[
            "llama-3.3-70b-versatile",
            "llama-3.1-8b-instant",
            "mixtral-8x7b-32768",
        ],
        index=["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "mixtral-8x7b-32768"]
              .index(saved.get("model", "llama-3.3-70b-versatile")),
        help="llama-3.3-70b: beste Qualität | llama-3.1-8b: schneller",
    )

    if groq_key and (groq_key != saved.get("groq_key") or model != saved.get("model")):
        save_config({"groq_key": groq_key, "model": model})

    st.markdown("---")
    if not groq_key:
        st.warning("⚠️ Noch kein API Key eingegeben.")
        st.markdown("[→ Kostenlos registrieren](https://console.groq.com)")
    else:
        st.success("✅ API Key gespeichert")

# ── Hauptbereich ──────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    start_date = st.date_input(
        "Von",
        value=date.today().replace(day=1),
        format="DD.MM.YYYY",
    )

with col2:
    end_date = st.date_input(
        "Bis",
        value=date.today(),
        format="DD.MM.YYYY",
    )

issue_number = st.text_input(
    "Issue-Nummer (optional)",
    value=saved.get("last_issue", ""),
    placeholder="z.B. 14",
)

st.divider()

# ── Start-Button ──────────────────────────────────────────────────────────────
start_disabled = not groq_key
if start_disabled:
    st.info("👈 Bitte zuerst den Groq API Key in der Seitenleiste eingeben.")

if st.button("🚀 Newsletter erstellen", type="primary",
             use_container_width=True, disabled=start_disabled):

    if start_date > end_date:
        st.error("Das Startdatum muss vor dem Enddatum liegen.")
        st.stop()

    issue_str = issue_number.strip() or "?"
    save_config({"groq_key": groq_key, "model": model, "last_issue": issue_str})

    import config
    config.GROQ_API_KEY = groq_key

    import summarizer
    summarizer.MODEL = model

    with st.status("⏳ Daten werden geladen...", expanded=True) as status:

        # ── Schritt 1: Artikel scrapen ────────────────────────────────────────
        st.write(f"📡 Suche Artikel: {start_date.strftime('%d.%m.%Y')} – {end_date.strftime('%d.%m.%Y')}")

        try:
            from scraper import fetch_articles_in_range
            articles = fetch_articles_in_range(start_date, end_date, verbose=False)
        except Exception as e:
            status.update(label="❌ Fehler beim Laden der Artikel", state="error")
            st.error(f"Fehler: {e}")
            st.stop()

        if not articles:
            status.update(label="Keine Artikel gefunden", state="error")
            st.warning(
                "Für diesen Zeitraum wurden keine Artikel gefunden. "
                "Bitte anderen Zeitraum wählen."
            )
            st.stop()

        st.write(f"✅ **{len(articles)} Artikel gefunden:**")
        for a in articles:
            st.write(f"  - {a.title[:75]}")

        # ── Schritt 2: Events scrapen ─────────────────────────────────────────
        # Upcoming events: vom Ende des Newsletter-Zeitraums, 1 Monat voraus
        import calendar as _cal
        _m = end_date.month - 1 + 1
        events_start = end_date
        events_end = end_date.replace(
            year=end_date.year + _m // 12,
            month=_m % 12 + 1,
            day=min(end_date.day, _cal.monthrange(end_date.year + _m // 12, _m % 12 + 1)[1]),
        )
        st.write(
            f"📅 Suche Upcoming Events: "
            f"{events_start.strftime('%d.%m.%Y')} – {events_end.strftime('%d.%m.%Y')}"
        )
        events = []
        events_truncated = False
        try:
            from event_scraper import fetch_events_in_range
            events = fetch_events_in_range(events_start, events_end, verbose=False)
            if len(events) > 10:
                events = events[:10]
                events_truncated = True
            if events:
                st.write(f"✅ **{len(events)} Events gefunden{' (auf 10 begrenzt)' if events_truncated else ''}:**")
                for ev in events:
                    date_str = ev.event_date.strftime("%d.%m.%Y")
                    st.write(f"  - {date_str}: {ev.title[:60]}")
            else:
                st.write("ℹ️ Keine Events für diesen Zeitraum gefunden.")
        except Exception as e:
            st.warning(f"Events konnten nicht geladen werden (Newsletter wird trotzdem erstellt): {e}")
            events = []

        # ── Schritt 3: Zusammenfassen ─────────────────────────────────────────
        st.write(f"🤖 Zusammenfassungen werden erstellt ({model})...")
        progress_bar = st.progress(0, text="Starte...")

        try:
            from summarizer import summarize_article, _build_client
            client = _build_client()
        except Exception as e:
            status.update(label="❌ API-Fehler", state="error")
            st.error(f"Verbindung zu Groq fehlgeschlagen: {e}")
            st.stop()

        summarized = []
        for i, article in enumerate(articles):
            progress_bar.progress(
                (i + 1) / len(articles),
                text=f"Artikel {i+1}/{len(articles)}: {article.title[:50]}..."
            )
            summary = summarize_article(article, client)
            summarized.append({
                "title": article.title,
                "author": article.author,
                "url": article.url,
                "published": article.published,
                "summary": summary,
                "image_url": article.image_url,
                "image_bytes": article.image_bytes,
            })

        progress_bar.empty()
        st.write("✅ Alle Zusammenfassungen fertig.")

        # ── Schritt 4: Events in Dicts umwandeln ──────────────────────────────
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
        ] if events is not None else []

        # ── Schritt 5: PPTX erstellen ─────────────────────────────────────────
        st.write("📄 PPTX wird generiert...")

        try:
            from pptx_builder import build_newsletter
            output_path = build_newsletter(
                articles=summarized,
                month=end_date.month,
                year=end_date.year,
                issue_number=issue_str,
                events=event_dicts,
                events_truncated=events_truncated,
            )
        except Exception as e:
            status.update(label="❌ Fehler bei PPTX-Erstellung", state="error")
            st.error(f"Fehler: {e}")
            st.stop()

        status.update(label="✅ Newsletter fertig!", state="complete", expanded=False)

    # ── Download ──────────────────────────────────────────────────────────────
    with open(output_path, "rb") as f:
        pptx_bytes = f.read()

    filename = os.path.basename(output_path)
    st.success(f"**{filename}** wurde erfolgreich erstellt!")

    st.download_button(
        label="⬇️ PPTX herunterladen",
        data=pptx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
        type="primary",
    )

    # ── Vorschau ──────────────────────────────────────────────────────────────
    with st.expander(f"📋 {len(summarized)} Artikel im Newsletter"):
        for a in summarized:
            st.markdown(f"**{a['title']}**")
            st.caption(
                f"{a['author']} · {a['published'].strftime('%d.%m.%Y')} "
                f"· [Originalartikel →]({a['url']})"
            )
            st.markdown(a["summary"])
            st.divider()

    if events:
        with st.expander(f"📅 {len(events)} Events im Newsletter"):
            for ev in events:
                date_str = ev.event_date.strftime("%d.%m.%Y")
                parts = [date_str]
                if ev.time_str:
                    parts.append(ev.time_str)
                if ev.location:
                    parts.append(ev.location)
                st.markdown(f"**{ev.title}**")
                st.caption("  ·  ".join(parts))
                if ev.description:
                    st.markdown(ev.description)
                if ev.url:
                    st.markdown(f"[→ Mehr Informationen]({ev.url})")
                st.divider()
