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

## Einfache Schritt-für-Schritt-Anleitung für Laien

### Voraussetzungen
- Python 3.9 oder höher ist installiert
- Git ist installiert

### Schritt 1: Projekt herunterladen
1. Öffnen Sie das Terminal.
2. Führen Sie folgenden Befehl aus:
   ```bash
   git clone git@github.com:roman6s/Z_Newsletter_Automation.git
   cd Z_Newsletter_Automation
   ```

### Schritt 2: Abhängigkeiten installieren
1. Geben Sie diesen Befehl ein:
   ```bash
   pip install -r requirements.txt
   ```

### Hinweis für macOS-Nutzer
Falls der Befehl `pip` nicht gefunden wird, verwenden Sie stattdessen `pip3`:
```bash
pip3 install -r requirements.txt
```

### Schritt 3: Anwendung starten
1. Starten Sie die Anwendung mit:
   ```bash
   python3 -m streamlit run app.py
   ```
2. Ein Browserfenster öffnet sich automatisch. Falls nicht, kopieren Sie den angezeigten Link (z. B. `http://localhost:8501`) und fügen Sie ihn in Ihren Browser ein.

### Hinweis für Streamlit
Falls der Befehl `streamlit` nicht gefunden wird, starten Sie die Anwendung mit folgendem Befehl:
```bash
python3 -m streamlit run app.py
```

### Schritt 4: Anwendung nutzen
1. Geben Sie Ihren Groq API Key in der Seitenleiste ein.
2. Wählen Sie den gewünschten Zeitraum aus.
3. Klicken Sie auf **Newsletter erstellen**.
4. Laden Sie die erstellte PPTX-Datei herunter.
