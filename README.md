# IBM Z Newsletter Automation

Erstellt automatisch einen IBM Z DACH Newsletter als PPTX aus den aktuellen Blog-Artikeln der [IBM Z DACH Community](https://community.ibm.com/community/user/groups/community-home/recent-community-blogs?CommunityKey=9a8b7fc3-b167-447a-8e14-adf93406eccc).

## Setup (einmalig)

**1. Repository klonen**
```bash
git clone git@github.com:roman6s/Z_Newsletter_Automation.git
cd Z_Newsletter_Automation
```

**2. Abhängigkeiten installieren**
```bash
pip install -r requirements.txt
```

**3. Groq API Key holen** (kostenlos)
→ https://console.groq.com → Registrieren → API Keys → Create API Key

## Starten

```bash
streamlit run app.py
```

Der Browser öffnet sich automatisch. Dann:
1. Groq API Key in der Seitenleiste eingeben
2. Zeitraum auswählen
3. **Newsletter erstellen** klicken
4. PPTX herunterladen

## Alternativ: Kommandozeile

```bash
export GROQ_API_KEY='gsk_...'
python main.py
```

## Projektstruktur

```
├── app.py                    # Streamlit Web-App (empfohlen)
├── main.py                   # Kommandozeilen-Alternative
├── scraper.py                # Holt Artikel vom IBM Community Blog
├── summarizer.py             # KI-Zusammenfassung via Groq
├── pptx_builder.py           # Befüllt das Newsletter-Template
├── config.py                 # Konfiguration
├── Oktober Newsletter.pptx   # Newsletter-Template
└── requirements.txt
```
