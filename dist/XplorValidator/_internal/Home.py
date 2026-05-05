"""
Xplor Data Migration Tools — Landing Page
==========================================
Run with:
    py -m streamlit run Home.py
"""

import streamlit as st

from shared.service_map import load_service_map_sidebar
from shared.styles import apply_styles

st.set_page_config(
    page_title="Xplor Data Migration Tools",
    page_icon="🏠",
    layout="wide",
)

apply_styles()

# ── Sidebar ───────────────────────────────────────────────────────────────────
load_service_map_sidebar()

with st.sidebar:
    st.markdown("## 📖 Getting Started")
    st.markdown("---")
    st.markdown("""
<div class="how-to-step">
  <div class="step-num">1</div>
  <div class="step-text"><strong>Upload serviceIDs.csv</strong><br>
  Upload your service ID mapping file in the sidebar above. It will be remembered as you move between tools.</div>
</div>

<div class="how-to-step">
  <div class="step-num">2</div>
  <div class="step-text"><strong>Choose a tool</strong><br>
  Select a tool from the left navigation panel below.</div>
</div>

<div class="how-to-step">
  <div class="step-num">3</div>
  <div class="step-text"><strong>Follow on-screen steps</strong><br>
  Each tool has its own upload prompts and run button.</div>
</div>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="tool-header">
    <h1>🚀 Xplor Data Migration Tools</h1>
    <p>A suite of tools to validate, clean, and prepare QikKids data for Xplor import.</p>
</div>
""", unsafe_allow_html=True)

# ── Reminder ──────────────────────────────────────────────────────────────────
if "service_map_bytes" not in st.session_state:
    st.info("👆 Upload your **serviceIDs.csv** in the sidebar before using any tool — it is required by every tool on this platform.")

st.markdown("### Available Tools")

# ── Tool cards ────────────────────────────────────────────────────────────────
TOOLS = [
    {
        "icon": "🔍",
        "name": "Xplor Validator",
        "page": "Xplor_Validator",
        "description": (
            "Validates your full migration CSV/XLSX against Xplor import rules. "
            "Flags errors, warnings, and auto-fixes, then saves split output CSVs and audit reports."
        ),
    },
    {
        "icon": "🏷️",
        "name": "Room & Fee Names",
        "page": "Room_Fee_Names",
        "description": (
            "Compares fee and room names used in QikKids bookings against those defined in Xplor. "
            "Produces an Excel mismatch report with fuzzy-match suggestions."
        ),
    },
    {
        "icon": "📅",
        "name": "Bookings Import",
        "page": "Bookings_Import",
        "description": (
            "Processes recurring and casual QikKids booking exports, removes duplicates and "
            "schedule overlaps, and produces split CSV files per service ready for Xplor import."
        ),
    },
    {
        "icon": "💳",
        "name": "Payment Plans",
        "page": "Payment_Plans",
        "description": (
            "Validates payment plan CSVs against Onboarding Tool error rules, auto-fixes date "
            "formats and weekday abbreviations, and splits output by service."
        ),
    },
    {
        "icon": "⚖️",
        "name": "Balance Adjustments",
        "page": "Balance_Adjustments",
        "description": (
            "Reads balance adjustment data, maps centre names to Xplor service names, removes "
            "demo accounts, and produces one styled XLSX file per centre using the upload template."
        ),
    },
]

for tool in TOOLS:
    st.markdown(f"""
<div class="tool-card">
    <h3><a href="/{tool['page']}" target="_self" style="color:#0f3460;text-decoration:none;">{tool['icon']} {tool['name']}</a></h3>
    <p>{tool['description']}</p>
</div>
""", unsafe_allow_html=True)

st.markdown("---")
st.caption("Xplor Data Migration Tools · Does not modify your original files.")
