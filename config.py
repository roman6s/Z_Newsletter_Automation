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

# Max articles per newsletter (template capacity: 4 slides × 2 articles)
MAX_ARTICLES = 8

# Summary language
SUMMARY_LANGUAGE = "Deutsch"

# Output directory
OUTPUT_DIR = "output"

# Template file
TEMPLATE_FILE = "/Users/romansivirin/Z_Newsletter_Automation/Oktober Newsletter.pptx"
