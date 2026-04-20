"""
IBM Z Newsletter Automation – Streamlit Frontend
Run with: streamlit run app.py
"""

import os
from datetime import date

import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="IBM Z Newsletter",
    page_icon="📰",
    layout="centered",
)

# ── Passwort-Schutz ───────────────────────────────────────────────────────────
APP_PASSWORD = st.secrets.get("APP_PASSWORD", os.environ.get("APP_PASSWORD", ""))

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("📰 IBM Z Newsletter")
    st.caption("IBM Z DACH Community – Newsletter Automation")
    st.divider()
    pwd = st.text_input("Passwort", type="password", placeholder="Passwort eingeben...")
    if st.button("Anmelden", type="primary", use_container_width=True):
        if APP_PASSWORD and pwd == APP_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Falsches Passwort.")
    st.stop()

# ── Session-State Defaults ────────────────────────────────────────────────────
# Alle Einstellungen leben nur in der Browser-Session (sicher, kein Server-File)
for key, default in {
    "provider":        "Groq",
    "api_key":         "",
    "model":           "llama-3.3-70b-versatile",
    "custom_base_url": "",
    "issue_number":    "",
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ── Provider-Konfiguration ────────────────────────────────────────────────────
PROVIDERS = {
    "Groq": {
        "base_url":    "https://api.groq.com/openai/v1",
        "placeholder": "gsk_...",
        "models":      ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "mixtral-8x7b-32768"],
        "key_link":    "https://console.groq.com",
        "key_label":   "console.groq.com",
        "hint":        "Kostenlos · schnell · empfohlen",
    },
    "OpenAI": {
        "base_url":    None,
        "placeholder": "sk-...",
        "models":      ["gpt-4o", "gpt-4o-mini", "gpt-3.5-turbo"],
        "key_link":    "https://platform.openai.com/api-keys",
        "key_label":   "platform.openai.com/api-keys",
        "hint":        "GPT-Modelle (kostenpflichtig)",
    },
    "Google Gemini": {
        "base_url":    "https://generativelanguage.googleapis.com/v1beta/openai/",
        "placeholder": "AIza...",
        "models":      ["gemini-2.0-flash", "gemini-1.5-pro", "gemini-1.5-flash"],
        "key_link":    "https://aistudio.google.com/apikey",
        "key_label":   "aistudio.google.com",
        "hint":        "Gemini-Modelle · kostenloses Kontingent verfügbar",
    },
    "Andere (OpenAI-kompatibel)": {
        "base_url":    None,
        "placeholder": "",
        "models":      [],
        "key_link":    None,
        "key_label":   None,
        "hint":        "Mistral, Together AI, Azure OpenAI, …",
    },
}

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("🤖 KI-Anbieter")

    provider_name = st.selectbox(
        "Anbieter",
        options=list(PROVIDERS.keys()),
        index=list(PROVIDERS.keys()).index(st.session_state.provider),
        label_visibility="collapsed",
        key="provider",
    )
    cfg = PROVIDERS[provider_name]
    st.caption(cfg["hint"])

    has_key = bool(st.session_state.api_key)
    with st.expander(
        "Noch keinen Key? Hier entlang →" if not has_key else "Key ändern",
        expanded=not has_key and provider_name == "Groq",
    ):
        if provider_name == "Groq":
            st.markdown(
                "**Einmalig, kostenlos, 2 Minuten:**\n\n"
                "1. [console.groq.com](https://console.groq.com) öffnen\n"
                "2. Account erstellen (Google-Login möglich)\n"
                "3. **API Keys → Create API Key**\n"
                "4. Key unten einfügen"
            )
        elif cfg["key_link"]:
            st.markdown(f"Key erstellen: [{cfg['key_label']}]({cfg['key_link']})")

    st.text_input(
        "API Key",
        type="password",
        placeholder=cfg["placeholder"] or "API Key...",
        label_visibility="collapsed",
        key="api_key",
    )

    if provider_name == "Andere (OpenAI-kompatibel)":
        st.text_input(
            "API Base URL",
            placeholder="https://api.example.com/v1",
            key="custom_base_url",
        )

    if st.session_state.api_key:
        st.success("✅ Key eingegeben")

    st.divider()
    st.header("⚙️ Modell")

    if cfg["models"]:
        # Modell zurücksetzen wenn es zum neuen Anbieter nicht passt
        if st.session_state.model not in cfg["models"]:
            st.session_state.model = cfg["models"][0]
        st.selectbox(
            "Modell",
            options=cfg["models"],
            label_visibility="collapsed",
            key="model",
        )
    else:
        st.text_input(
            "Modell (Name eingeben)",
            placeholder="z.B. mistral-large-latest",
            key="model",
        )

    st.markdown("---")
    if st.button("Abmelden", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

# ── Hauptbereich ──────────────────────────────────────────────────────────────
st.title("📰 IBM Z Newsletter Automation")
st.caption("Erstellt automatisch einen Newsletter aus den aktuellen IBM Z DACH Blog-Artikeln.")
st.divider()

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Von", value=date.today().replace(day=1), format="DD.MM.YYYY")
with col2:
    end_date = st.date_input("Bis", value=date.today(), format="DD.MM.YYYY")

st.text_input("Issue-Nummer (optional)", placeholder="z.B. 14", key="issue_number")

st.divider()

# ── Start-Button ──────────────────────────────────────────────────────────────
api_key  = st.session_state.api_key
model    = st.session_state.model
base_url = (st.session_state.custom_base_url
            if provider_name == "Andere (OpenAI-kompatibel)"
            else cfg["base_url"] or "")
ready    = bool(api_key) and bool(model) and (
               bool(base_url) if provider_name == "Andere (OpenAI-kompatibel)" else True
           )

if not ready:
    st.info("👈 Bitte zuerst API Key und Modell in der Seitenleiste eingeben.")

if st.button("🚀 Newsletter erstellen", type="primary",
             use_container_width=True, disabled=not ready):

    if start_date > end_date:
        st.error("Das Startdatum muss vor dem Enddatum liegen.")
        st.stop()

    issue_str = st.session_state.issue_number.strip() or "?"

    import summarizer
    summarizer.API_KEY  = api_key
    summarizer.BASE_URL = base_url
    summarizer.MODEL    = model

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
            st.warning("Für diesen Zeitraum wurden keine Artikel gefunden. Bitte anderen Zeitraum wählen.")
            st.stop()

        st.write(f"✅ **{len(articles)} Artikel gefunden:**")
        for a in articles:
            st.write(f"  - {a.title[:75]}")

        # ── Schritt 2: Events scrapen ─────────────────────────────────────────
        import calendar as _cal
        _m = end_date.month - 1 + 1
        events_start = end_date
        events_end = end_date.replace(
            year=end_date.year + _m // 12,
            month=_m % 12 + 1,
            day=min(end_date.day, _cal.monthrange(end_date.year + _m // 12, _m % 12 + 1)[1]),
        )
        st.write(f"📅 Suche Upcoming Events: {events_start.strftime('%d.%m.%Y')} – {events_end.strftime('%d.%m.%Y')}")
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
                    st.write(f"  - {ev.event_date.strftime('%d.%m.%Y')}: {ev.title[:60]}")
            else:
                st.write("ℹ️ Keine Events für diesen Zeitraum gefunden.")
        except Exception as e:
            st.warning(f"Events konnten nicht geladen werden (Newsletter wird trotzdem erstellt): {e}")

        # ── Schritt 3: Zusammenfassen ─────────────────────────────────────────
        st.write(f"🤖 Zusammenfassungen werden erstellt ({model})...")
        progress_bar = st.progress(0, text="Starte...")
        try:
            from summarizer import summarize_article, _build_client
            client = _build_client()
        except Exception as e:
            status.update(label="❌ API-Fehler", state="error")
            st.error(f"Verbindung zur KI fehlgeschlagen: {e}")
            st.stop()

        summarized = []
        for i, article in enumerate(articles):
            progress_bar.progress(
                (i + 1) / len(articles),
                text=f"Artikel {i+1}/{len(articles)}: {article.title[:50]}..."
            )
            summary = summarize_article(article, client)
            summarized.append({
                "title":       article.title,
                "author":      article.author,
                "url":         article.url,
                "published":   article.published,
                "summary":     summary,
                "image_url":   article.image_url,
                "image_bytes": article.image_bytes,
            })

        progress_bar.empty()
        st.write("✅ Alle Zusammenfassungen fertig.")

        # ── Schritt 4: Events in Dicts umwandeln ──────────────────────────────
        event_dicts = [
            {
                "title":       ev.title,
                "event_date":  ev.event_date,
                "time_str":    ev.time_str,
                "location":    ev.location,
                "description": ev.description,
                "url":         ev.url,
            }
            for ev in events
        ] if events else []

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
            st.caption(f"{a['author']} · {a['published'].strftime('%d.%m.%Y')} · [Originalartikel →]({a['url']})")
            st.markdown(a["summary"])
            st.divider()

    if events:
        with st.expander(f"📅 {len(events)} Events im Newsletter"):
            for ev in events:
                parts = [ev.event_date.strftime("%d.%m.%Y")]
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
