"""
Configuration for IBM Z Newsletter Automation.

API Key Setup:
  Option A (empfohlen): In der Streamlit-App unter "Einstellungen" eingeben.
  Option B: Als Umgebungsvariable setzen:
            export GROQ_API_KEY='gsk_...'
  Option C: Direkt hier eintragen (nicht für geteilte Repos geeignet):
            GROQ_API_KEY = "gsk_..."

Kostenlosen Groq Key erstellen: https://console.groq.com
"""

import os

# Groq API Key
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")

# Blog URL
BLOG_URL = (
    "https://community.ibm.com/community/user/groups/community-home/"
    "recent-community-blogs?CommunityKey=9a8b7fc3-b167-447a-8e14-adf93406eccc"
)

# Events URL
EVENTS_URL = (
    "https://community.ibm.com/community/user/events/calendar"
    "?CommunityKey=9a8b7fc3-b167-447a-8e14-adf93406eccc"
)

# Summary language
SUMMARY_LANGUAGE = "Deutsch"

# Output directory (relative to this file)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# Template file (relative to this file – works regardless of working directory)
TEMPLATE_FILE = os.path.join(BASE_DIR, "Oktober Newsletter.pptx")
