import streamlit as st
import pandas as pd
import numpy as np
import json
import time
from datetime import datetime, timedelta

# Try loading gspread for sheet sync
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# Try loading Google GenAI
try:
    from google import genai
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False

# ==========================================
# 1. PAGE CONFIGURATION & UI STYLING
# ==========================================
st.set_page_config(page_title="D2C Scaler | Performance Marketing Command Center", layout="wide")

# Custom CSS for Premium Dark Glassmorphism Aesthetics
st.markdown("""
<style>
    /* Global Styles */
    .stApp {
        background-color: #030712;
        color: #f9fafb;
    }
    
    /* Title and headers */
    h1, h2, h3 {
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 700 !important;
    }
    
    /* KPI metric cards style */
    .kpi-container {
        display: flex;
        justify-content: space-between;
        gap: 15px;
        margin-bottom: 25px;
        flex-wrap: wrap;
    }
    .kpi-card {
        flex: 1;
        min-width: 220px;
        background-color: rgba(17, 24, 39, 0.65);
        border: 1px solid rgba(30, 41, 59, 0.7);
        border-radius: 14px;
        padding: 20px;
        text-align: left;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        backdrop-filter: blur(16px);
    }
    .kpi-card-accent {
        border-color: rgba(99, 102, 241, 0.5) !important;
        background: radial-gradient(100% 100% at 100% 0%, rgba(99, 102, 241, 0.08) 0%, transparent 100%), rgba(17, 24, 39, 0.65) !important;
    }
    .kpi-title {
        color: #9ca3af;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 8px;
    }
    .kpi-value {
        color: #f9fafb;
        font-size: 26px;
        font-weight: 700;
        font-family: 'Space Grotesk', sans-serif;
        margin-bottom: 4px;
    }
    .kpi-delta {
        font-size: 12px;
        font-weight: 600;
        display: flex;
        align-items: center;
        gap: 4px;
    }
    .delta-up { color: #10b981; }
    .delta-down { color: #ef4444; }
    .delta-neutral { color: #9ca3af; }
    
    /* Code and preformatted blocks */
    code {
        background-color: #050811 !important;
        color: #a5b4fc !important;
        padding: 2px 6px !important;
        border-radius: 4px !important;
        font-family: 'Space Grotesk', monospace !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. SESSION STATES AND SWIPE PRELOADS
# ==========================================
# 14 Days of Realistic D2C Data pre-loaded
if "raw_data" not in st.session_state:
    st.session_state.raw_data = pd.DataFrame([
        # HairSerum - Meta Ads (Scaling Case)
        { "Date": "2026-05-24", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 800, "Impressions": 20000, "Clicks": 650, "Purchases": 4, "Revenue": 3600, "Three_Sec_Views": 8000, "Landing_Page_Views": 550, "Add_To_Cart": 45 },
        { "Date": "2026-05-25", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 900, "Impressions": 22000, "Clicks": 700, "Purchases": 5, "Revenue": 4500, "Three_Sec_Views": 9000, "Landing_Page_Views": 600, "Add_To_Cart": 55 },
        { "Date": "2026-05-26", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1000, "Impressions": 25000, "Clicks": 800, "Purchases": 5, "Revenue": 4500, "Three_Sec_Views": 10500, "Landing_Page_Views": 700, "Add_To_Cart": 60 },
        { "Date": "2026-05-27", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1100, "Impressions": 26000, "Clicks": 820, "Purchases": 6, "Revenue": 5400, "Three_Sec_Views": 11000, "Landing_Page_Views": 720, "Add_To_Cart": 62 },
        { "Date": "2026-05-28", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1200, "Impressions": 29000, "Clicks": 900, "Purchases": 7, "Revenue": 6300, "Three_Sec_Views": 12500, "Landing_Page_Views": 810, "Add_To_Cart": 75 },
        { "Date": "2026-05-29", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1200, "Impressions": 28000, "Clicks": 890, "Purchases": 6, "Revenue": 5400, "Three_Sec_Views": 12000, "Landing_Page_Views": 780, "Add_To_Cart": 70 },
        { "Date": "2026-05-30", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1200, "Impressions": 27000, "Clicks": 850, "Purchases": 7, "Revenue": 6300, "Three_Sec_Views": 11500, "Landing_Page_Views": 760, "Add_To_Cart": 72 },
        { "Date": "2026-05-31", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1200, "Impressions": 28000, "Clicks": 870, "Purchases": 8, "Revenue": 7200, "Three_Sec_Views": 12000, "Landing_Page_Views": 770, "Add_To_Cart": 74 },
        { "Date": "2026-06-01", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1300, "Impressions": 30000, "Clicks": 920, "Purchases": 9, "Revenue": 8100, "Three_Sec_Views": 13000, "Landing_Page_Views": 830, "Add_To_Cart": 80 },
        { "Date": "2026-06-02", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1400, "Impressions": 32000, "Clicks": 950, "Purchases": 10, "Revenue": 9000, "Three_Sec_Views": 14000, "Landing_Page_Views": 870, "Add_To_Cart": 90 },
        { "Date": "2026-06-03", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1500, "Impressions": 33000, "Clicks": 1020, "Purchases": 11, "Revenue": 9900, "Three_Sec_Views": 15000, "Landing_Page_Views": 920, "Add_To_Cart": 100 },
        # Last 3 days scaling
        { "Date": "2026-06-04", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 1800, "Impressions": 40000, "Clicks": 1300, "Purchases": 15, "Revenue": 13500, "Three_Sec_Views": 18000, "Landing_Page_Views": 1100, "Add_To_Cart": 130 },
        { "Date": "2026-06-05", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 2000, "Impressions": 45000, "Clicks": 1500, "Purchases": 18, "Revenue": 16200, "Three_Sec_Views": 20000, "Landing_Page_Views": 1300, "Add_To_Cart": 150 },
        { "Date": "2026-06-06", "Platform": "Meta", "Product": "HairSerum", "Ad_Name": "Meta_Serum_UGC_01", "Spend": 2200, "Impressions": 50000, "Clicks": 1700, "Purchases": 20, "Revenue": 18000, "Three_Sec_Views": 23000, "Landing_Page_Views": 1500, "Add_To_Cart": 180 },

        # FaceWash - Meta Ads (Fatigue & Pause Case)
        { "Date": "2026-05-24", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1000, "Impressions": 20000, "Clicks": 600, "Purchases": 6, "Revenue": 3000, "Three_Sec_Views": 4000, "Landing_Page_Views": 500, "Add_To_Cart": 40 },
        { "Date": "2026-05-25", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1000, "Impressions": 21000, "Clicks": 590, "Purchases": 5, "Revenue": 2500, "Three_Sec_Views": 4200, "Landing_Page_Views": 490, "Add_To_Cart": 35 },
        { "Date": "2026-05-26", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1000, "Impressions": 19000, "Clicks": 580, "Purchases": 5, "Revenue": 2500, "Three_Sec_Views": 3800, "Landing_Page_Views": 480, "Add_To_Cart": 32 },
        { "Date": "2026-05-27", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1200, "Impressions": 24000, "Clicks": 650, "Purchases": 6, "Revenue": 3000, "Three_Sec_Views": 4800, "Landing_Page_Views": 550, "Add_To_Cart": 38 },
        { "Date": "2026-05-28", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1200, "Impressions": 23000, "Clicks": 610, "Purchases": 5, "Revenue": 2500, "Three_Sec_Views": 4600, "Landing_Page_Views": 510, "Add_To_Cart": 34 },
        { "Date": "2026-05-29", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1200, "Impressions": 22000, "Clicks": 600, "Purchases": 4, "Revenue": 2000, "Three_Sec_Views": 4400, "Landing_Page_Views": 500, "Add_To_Cart": 30 },
        { "Date": "2026-05-30", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1200, "Impressions": 23000, "Clicks": 620, "Purchases": 5, "Revenue": 2500, "Three_Sec_Views": 4500, "Landing_Page_Views": 520, "Add_To_Cart": 36 },
        { "Date": "2026-05-31", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1500, "Impressions": 30000, "Clicks": 750, "Purchases": 6, "Revenue": 3000, "Three_Sec_Views": 5500, "Landing_Page_Views": 650, "Add_To_Cart": 42 },
        { "Date": "2026-06-01", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1500, "Impressions": 29000, "Clicks": 710, "Purchases": 5, "Revenue": 2500, "Three_Sec_Views": 5200, "Landing_Page_Views": 610, "Add_To_Cart": 38 },
        { "Date": "2026-06-02", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1500, "Impressions": 31000, "Clicks": 730, "Purchases": 6, "Revenue": 3000, "Three_Sec_Views": 5600, "Landing_Page_Views": 630, "Add_To_Cart": 40 },
        { "Date": "2026-06-03", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 1500, "Impressions": 30000, "Clicks": 680, "Purchases": 4, "Revenue": 2000, "Three_Sec_Views": 5000, "Landing_Page_Views": 580, "Add_To_Cart": 33 },
        # Last 3 days fatigue
        { "Date": "2026-06-04", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 2000, "Impressions": 38000, "Clicks": 450, "Purchases": 2, "Revenue": 1000, "Three_Sec_Views": 4500, "Landing_Page_Views": 380, "Add_To_Cart": 12 },
        { "Date": "2026-06-05", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 2000, "Impressions": 39000, "Clicks": 400, "Purchases": 1, "Revenue": 500, "Three_Sec_Views": 4000, "Landing_Page_Views": 340, "Add_To_Cart": 8 },
        { "Date": "2026-06-06", "Platform": "Meta", "Product": "FaceWash", "Ad_Name": "Meta_Wash_Static_02", "Spend": 2200, "Impressions": 42000, "Clicks": 380, "Purchases": 1, "Revenue": 500, "Three_Sec_Views": 3900, "Landing_Page_Views": 310, "Add_To_Cart": 5 },

        # Sunscreen - Google Ads (Watch/Hold Case)
        { "Date": "2026-05-24", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 400, "Impressions": 5000, "Clicks": 300, "Purchases": 1, "Revenue": 800, "Three_Sec_Views": 0, "Landing_Page_Views": 270, "Add_To_Cart": 20 },
        { "Date": "2026-05-25", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 400, "Impressions": 5200, "Clicks": 310, "Purchases": 2, "Revenue": 1600, "Three_Sec_Views": 0, "Landing_Page_Views": 280, "Add_To_Cart": 22 },
        { "Date": "2026-05-26", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 450, "Impressions": 5500, "Clicks": 340, "Purchases": 1, "Revenue": 800, "Three_Sec_Views": 0, "Landing_Page_Views": 300, "Add_To_Cart": 24 },
        { "Date": "2026-05-27", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 450, "Impressions": 5300, "Clicks": 320, "Purchases": 2, "Revenue": 1600, "Three_Sec_Views": 0, "Landing_Page_Views": 290, "Add_To_Cart": 23 },
        { "Date": "2026-05-28", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 500, "Impressions": 6000, "Clicks": 380, "Purchases": 2, "Revenue": 1600, "Three_Sec_Views": 0, "Landing_Page_Views": 340, "Add_To_Cart": 28 },
        { "Date": "2026-05-29", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 500, "Impressions": 5800, "Clicks": 360, "Purchases": 1, "Revenue": 800, "Three_Sec_Views": 0, "Landing_Page_Views": 320, "Add_To_Cart": 26 },
        { "Date": "2026-05-30", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 500, "Impressions": 5900, "Clicks": 370, "Purchases": 2, "Revenue": 1600, "Three_Sec_Views": 0, "Landing_Page_Views": 330, "Add_To_Cart": 27 },
        { "Date": "2026-05-31", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 500, "Impressions": 6200, "Clicks": 400, "Purchases": 2, "Revenue": 1600, "Three_Sec_Views": 0, "Landing_Page_Views": 360, "Add_To_Cart": 30 },
        { "Date": "2026-06-01", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 500, "Impressions": 6100, "Clicks": 390, "Purchases": 1, "Revenue": 800, "Three_Sec_Views": 0, "Landing_Page_Views": 350, "Add_To_Cart": 29 },
        { "Date": "2026-06-02", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 500, "Impressions": 6300, "Clicks": 410, "Purchases": 2, "Revenue": 1600, "Three_Sec_Views": 0, "Landing_Page_Views": 370, "Add_To_Cart": 32 },
        { "Date": "2026-06-03", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 500, "Impressions": 6000, "Clicks": 380, "Purchases": 2, "Revenue": 1600, "Three_Sec_Views": 0, "Landing_Page_Views": 340, "Add_To_Cart": 31 },
        # Last 3 days watch/hold
        { "Date": "2026-06-04", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 600, "Impressions": 7200, "Clicks": 480, "Purchases": 2, "Revenue": 1600, "Three_Sec_Views": 0, "Landing_Page_Views": 430, "Add_To_Cart": 35 },
        { "Date": "2026-06-05", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 600, "Impressions": 7100, "Clicks": 470, "Purchases": 2, "Revenue": 1600, "Three_Sec_Views": 0, "Landing_Page_Views": 420, "Add_To_Cart": 34 },
        { "Date": "2026-06-06", "Platform": "Google", "Product": "Sunscreen", "Ad_Name": "G_Search_Sunscreen_Exact", "Spend": 600, "Impressions": 7300, "Clicks": 490, "Purchases": 3, "Revenue": 2400, "Three_Sec_Views": 0, "Landing_Page_Views": 440, "Add_To_Cart": 38 },

        # HairSerum - Google Ads (Scaling Case)
        { "Date": "2026-05-24", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 500, "Impressions": 8000, "Clicks": 400, "Purchases": 2, "Revenue": 1800, "Three_Sec_Views": 0, "Landing_Page_Views": 360, "Add_To_Cart": 30 },
        { "Date": "2026-05-25", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 500, "Impressions": 8200, "Clicks": 410, "Purchases": 2, "Revenue": 1800, "Three_Sec_Views": 0, "Landing_Page_Views": 370, "Add_To_Cart": 32 },
        { "Date": "2026-05-26", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 500, "Impressions": 8500, "Clicks": 440, "Purchases": 3, "Revenue": 2700, "Three_Sec_Views": 0, "Landing_Page_Views": 400, "Add_To_Cart": 35 },
        { "Date": "2026-05-27", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 500, "Impressions": 8300, "Clicks": 420, "Purchases": 2, "Revenue": 1800, "Three_Sec_Views": 0, "Landing_Page_Views": 380, "Add_To_Cart": 31 },
        { "Date": "2026-05-28", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 600, "Impressions": 10000, "Clicks": 520, "Purchases": 3, "Revenue": 2700, "Three_Sec_Views": 0, "Landing_Page_Views": 470, "Add_To_Cart": 40 },
        { "Date": "2026-05-29", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 600, "Impressions": 9800, "Clicks": 500, "Purchases": 3, "Revenue": 2700, "Three_Sec_Views": 0, "Landing_Page_Views": 450, "Add_To_Cart": 38 },
        { "Date": "2026-05-30", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 600, "Impressions": 9900, "Clicks": 510, "Purchases": 4, "Revenue": 3600, "Three_Sec_Views": 0, "Landing_Page_Views": 460, "Add_To_Cart": 42 },
        { "Date": "2026-05-31", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 600, "Impressions": 10200, "Clicks": 540, "Purchases": 3, "Revenue": 2700, "Three_Sec_Views": 0, "Landing_Page_Views": 490, "Add_To_Cart": 41 },
        { "Date": "2026-06-01", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 600, "Impressions": 10100, "Clicks": 530, "Purchases": 4, "Revenue": 3600, "Three_Sec_Views": 0, "Landing_Page_Views": 480, "Add_To_Cart": 43 },
        { "Date": "2026-06-02", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 600, "Impressions": 10300, "Clicks": 550, "Purchases": 4, "Revenue": 3600, "Three_Sec_Views": 0, "Landing_Page_Views": 500, "Add_To_Cart": 45 },
        { "Date": "2026-06-03", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 600, "Impressions": 10000, "Clicks": 520, "Purchases": 3, "Revenue": 2700, "Three_Sec_Views": 0, "Landing_Page_Views": 470, "Add_To_Cart": 42 },
        # Last 3 days scaling
        { "Date": "2026-06-04", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 800, "Impressions": 13000, "Clicks": 720, "Purchases": 6, "Revenue": 5400, "Three_Sec_Views": 0, "Landing_Page_Views": 650, "Add_To_Cart": 60 },
        { "Date": "2026-06-05", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 800, "Impressions": 12800, "Clicks": 710, "Purchases": 5, "Revenue": 4500, "Three_Sec_Views": 0, "Landing_Page_Views": 640, "Add_To_Cart": 55 },
        { "Date": "2026-06-06", "Platform": "Google", "Product": "HairSerum", "Ad_Name": "G_PMax_Serum_01", "Spend": 1000, "Impressions": 16000, "Clicks": 900, "Purchases": 8, "Revenue": 7200, "Three_Sec_Views": 0, "Landing_Page_Views": 810, "Add_To_Cart": 75 }
    ])

if "competitor_ads" not in st.session_state:
    st.session_state.competitor_ads = [
        {
            "Name": "Kosa Labs",
            "Platform": "Meta",
            "Hook": "Us vs Them (Comparison Table)",
            "Link": "https://www.facebook.com/ads/library/?id=87452391038",
            "Copy": "Why 94% of women are dumping normal chemical serums for our peptide active formulation. Clear comparison breakdown.",
            "Insights": "Match their comparison tables by creating a dedicated comparison visual comparing our Serum's concentration levels vs generic competitors."
        },
        {
            "Name": "SkinGlow",
            "Platform": "Google",
            "Hook": "Three Reasons Why (Feature Callout)",
            "Link": "https://adstransparency.google.com/?referrer=skinglow",
            "Copy": "Doctor recommended face wash. Three reasons why standard sulfate cleansers destroy your lipid barrier.",
            "Insights": "We should spin up Google Search ads targeting keywords like 'dermatologist approved face wash' directing to an educational landing page showing our sulfate-free tests."
        }
    ]

if "chat_history" not in st.session_state:
    st.session_state.chat_history = [
        {"role": "agent", "content": "👋 Hello! I am your D2C Scaler AI Agent. I have scanned the loaded account metrics. Ask me things like:\n\n• 'Which SKU should I scale immediately?'\n• 'Give me a breakdown of creative fatigue alerts'\n• 'Draft a Facebook Us vs Them ad copy variant based on competitor hooks'"}
    ]

# ==========================================
# 3. SIDEBAR CONFIGURATION
# ==========================================
st.sidebar.title("🚀 Navigation & Configuration")

# Section 1: Ingest Data
st.sidebar.header("1. Ingest Daily Data")
sheet_url = st.sidebar.text_input("Google Sheets Edit Link", value=st.session_state.get("sheet_url", ""))

if st.sidebar.button("🔄 Sync Google Sheet"):
    if sheet_url:
        st.session_state.sheet_url = sheet_url
        try:
            # Parse Spreadsheet ID
            import re
            sheet_id_match = re.search(r"/d/([a-zA-Z0-9-_]+)", sheet_url)
            if sheet_id_match:
                sheet_id = sheet_id_match.group(1)
                viz_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=RAW_DATA"
                
                # Fetch CSV live
                import urllib.request
                import io
                response = urllib.request.urlopen(viz_url)
                csv_bytes = response.read()
                df = pd.read_csv(io.BytesIO(csv_bytes))
                
                if len(df) > 0:
                    st.session_state.raw_data = df
                    st.sidebar.success(f"Synced {len(df)} records live!")
                    st.rerun()
                else:
                    st.sidebar.error("Parsed 0 rows from sheet. Check tab name RAW_DATA.")
            else:
                st.sidebar.error("Invalid URL format.")
        except Exception as e:
            st.sidebar.error(f"Sync failed: {str(e)}. Make sure sheet is shared as 'Anyone with link can view'.")
    else:
        st.sidebar.warning("Please input a sheet URL.")

st.sidebar.markdown("**— OR —**")
csv_paste = st.sidebar.text_area("Paste Excel/CSV/TSV", height=100, placeholder="Paste headers + rows here...")

if st.sidebar.button("📥 Ingest Pasted Data"):
    if csv_paste.strip():
        try:
            import io
            # Detect separator (tab for Excel/Sheets paste, comma for CSV)
            sep = '\t' if '\t' in csv_paste else ','
            df = pd.read_csv(io.StringIO(csv_paste), sep=sep)
            
            if len(df) > 0:
                # Align columns and types
                df.columns = [c.strip() for c in df.columns]
                
                # Clean currency symbols
                for col in df.columns:
                    if df[col].dtype == object:
                        try:
                            # Try to strip currency characters
                            clean_col = df[col].astype(str).str.replace('₹', '').str.replace(',', '').str.strip()
                            df[col] = pd.to_numeric(clean_col)
                        except:
                            pass
                
                # Save to session
                st.session_state.raw_data = df
                st.sidebar.success(f"Successfully ingested {len(df)} rows!")
                st.rerun()
            else:
                st.sidebar.error("Could not parse data rows.")
        except Exception as e:
            st.sidebar.error(f"Parsing failed: {str(e)}")
    else:
        st.sidebar.warning("Paste data before submitting.")

# Section 2: AI Credentials
st.sidebar.header("2. AI Credentials")
api_key = st.sidebar.text_input("Gemini API Key (Optional)", type="password", value=st.session_state.get("api_key", ""))
if api_key:
    st.session_state.api_key = api_key

# Section 3: Global Target Configs
st.sidebar.header("3. Global Targets")
target_roas = st.sidebar.slider("Target ROAS", 0.5, 5.0, 2.0, 0.1)
target_cpa = st.sidebar.number_input("Target CPA (₹)", min_value=100, max_value=5000, value=500, step=50)

# API Status lights
st.sidebar.markdown("---")
st.sidebar.markdown("**System Connection States:**")
st.sidebar.markdown(f"🟢 Google Sheet Ingestion: Active")
if api_key:
    st.sidebar.markdown("🟢 Gemini API: Direct Connected")
else:
    st.sidebar.markdown("🟡 Gemini API: Local Heuristic Fallback")

# ==========================================
# 4. DATA COMPILATION PIPELINE
# ==========================================
df = st.session_state.raw_data.copy()

# Ensure numeric columns are cleaned and loaded correctly
num_cols = ['Spend', 'Impressions', 'Clicks', 'Purchases', 'Revenue', 'Three_Sec_Views', 'Landing_Page_Views', 'Add_To_Cart']
for col in num_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace('₹', '').str.replace(',', '').str.strip(), errors='coerce').fillna(0)
    else:
        df[col] = 0

# Overall calculations
total_spend = df['Spend'].sum()
total_revenue = df['Revenue'].sum()
blended_roas = total_revenue / total_spend if total_spend > 0 else 0
total_purchases = df['Purchases'].sum()
blended_cpa = total_spend / total_purchases if total_purchases > 0 else 0
blended_ctr = (df['Clicks'].sum() / df['Impressions'].sum() * 100) if df['Impressions'].sum() > 0 else 0
blended_cpc = total_spend / df['Clicks'].sum() if df['Clicks'].sum() > 0 else 0

# Trend mapping (3-day vs preceding 7-day average)
df['Date'] = pd.to_datetime(df['Date'])
max_date = df['Date'].max()
recent_dates = pd.date_range(end=max_date, periods=3, freq='D')
baseline_dates = pd.date_range(end=max_date - timedelta(days=3), periods=7, freq='D')

recent_df = df[df['Date'].isin(recent_dates)]
baseline_df = df[df['Date'].isin(baseline_dates)]

# Scale factor to align 3-day spend against a 7-day baseline
scale_factor = 3 / 7
r_spend = recent_df['Spend'].sum()
r_rev = recent_df['Revenue'].sum()
r_purch = recent_df['Purchases'].sum()

b_spend = baseline_df['Spend'].sum()
b_rev = baseline_df['Revenue'].sum()
b_purch = baseline_df['Purchases'].sum()

scaled_b_spend = b_spend * scale_factor
scaled_b_rev = b_rev * scale_factor

recent_roas = r_rev / r_spend if r_spend > 0 else 0
base_roas = b_rev / b_spend if b_spend > 0 else 0
recent_cpa = r_spend / r_purch if r_purch > 0 else 0
base_cpa = b_spend / b_purch if b_purch > 0 else 0

spend_delta = ((r_spend - scaled_b_spend) / scaled_b_spend * 100) if scaled_b_spend > 0 else 0
rev_delta = ((r_rev - scaled_b_rev) / scaled_b_rev * 100) if scaled_b_rev > 0 else 0
roas_delta = ((recent_roas - base_roas) / base_roas * 100) if base_roas > 0 else 0
cpa_delta = ((recent_cpa - base_cpa) / base_cpa * 100) if base_cpa > 0 else 0

# ==========================================
# 5. TABBED INTERFACE ARCHITECTURE
# ==========================================
tab_docs, tab_dashboard, tab_products, tab_intelligence, tab_competitors, tab_agent = st.tabs([
    "📖 Manual",
    "📊 Executive Summary",
    "🛍️ Product Breakdown",
    "⚡ Intelligence Mode",
    "🕵️ Competitor Spy",
    "🤖 Growth Agent"
])

# --- TAB 1: MANUAL ---
with tab_docs:
    st.header("Operator's Manual & Execution Guide")
    st.markdown(f"""
    Welcome to the **D2C Performance Scaler Dashboard**. This system monitors creative efficiency and ad spend, processing raw campaign records to suggest optimization protocols.
    
    ### 🚀 Setup Instructions
    
    1. **Format Your Columns:** Whether syncing from Google Sheets or using the clipboard copy-paste tool, your table must follow these headers exactly:
    """)
    
    headers_txt = "Date,Platform,Product,Ad_Name,Spend,Impressions,Clicks,Purchases,Revenue,Three_Sec_Views,Landing_Page_Views,Add_To_Cart"
    st.code(headers_txt)
    
    st.markdown("""
    2. **Connect Google Sheet:**
       - Share your Sheet: **File -> Share -> Anyone with link can view**.
       - Ensure your tab is named exactly **`RAW_DATA`**.
       - Paste the link in the sidebar and sync.
       
    3. **Clipboard Paste (Free & Private):**
       - Simply copy your worksheet cells (Ctrl+C / Cmd+C).
       - Paste them into the sidebar text area and click **Ingest Pasted Data**.
       
    ### 📐 Trend Calculation Logic (3-Day vs 7-Day)
    Ad performance fluctuates. The **Intelligence Mode** uses the past 14 days of data to compute optimization verdicts:
    - **SCALE BUDGET (🚀):** Recent 3d ROAS is above target, recent 3d CPA is below target, and click-through rates (CTR) remain stable compared to baseline.
    - **PAUSE CREATIVE (🛑):** Recent ROAS is critically low (< 1.3x) or spend has exceeded 2.5x Target CPA without converting a single sale.
    - **CREATIVE FATIGUE (⚠️):** Recent CTR has dropped by >25% and CPA has increased by >20% compared to baseline average (indicating ad burnout).
    """)

# --- TAB 2: EXECUTIVE SUMMARY ---
with tab_dashboard:
    st.header("Executive Summary Dashboard")
    
    # Custom HTML metrics containers
    st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-card kpi-card-accent">
            <div class="kpi-title">💵 Blended Spend (3d)</div>
            <div class="kpi-value">₹{int(r_spend):,}</div>
            <div class="kpi-delta {'delta-up' if spend_delta >= 0 else 'delta-down'}">
                {'▲' if spend_delta >= 0 else '▼'} {abs(spend_delta):.1f}% vs baseline
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-title">🛍️ Blended Revenue (3d)</div>
            <div class="kpi-value">₹{int(r_rev):,}</div>
            <div class="kpi-delta {'delta-up' if rev_delta >= 0 else 'delta-down'}">
                {'▲' if rev_delta >= 0 else '▼'} {abs(rev_delta):.1f}% vs baseline
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-title">🎯 Blended ROAS (3d)</div>
            <div class="kpi-value">{recent_roas:.2f}x</div>
            <div class="kpi-delta {'delta-up' if roas_delta >= 0 else 'delta-down'}">
                {'▲' if roas_delta >= 0 else '▼'} {abs(roas_delta):.1f}% vs baseline
            </div>
        </div>
        <div class="kpi-card">
            <div class="kpi-title">🏷️ Blended CPA (CAC)</div>
            <div class="kpi-value">₹{int(recent_cpa):,}</div>
            <div class="kpi-delta {'delta-up' if cpa_delta <= 0 else 'delta-down'}">
                {'▼' if cpa_delta <= 0 else '▲'} {abs(cpa_delta):.1f}% vs baseline
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Spend Distribution by Platform")
        plat_spend = df.groupby('Platform')['Spend'].sum()
        st.bar_chart(plat_spend)
        
    with col2:
        st.subheader("Revenue Breakdown by Platform")
        plat_rev = df.groupby('Platform')['Revenue'].sum()
        st.bar_chart(plat_rev)

    st.subheader("Platform Summary Scorecard")
    plat_grouped = df.groupby('Platform').agg({
        'Spend': 'sum',
        'Revenue': 'sum',
        'Purchases': 'sum',
        'Clicks': 'sum',
        'Impressions': 'sum'
    }).reset_index()
    
    plat_grouped['ROAS'] = plat_grouped['Revenue'] / plat_grouped['Spend']
    plat_grouped['CPA'] = plat_grouped['Spend'] / plat_grouped['Purchases']
    plat_grouped['CTR'] = (plat_grouped['Clicks'] / plat_grouped['Impressions']) * 100
    plat_grouped['CPC'] = plat_grouped['Spend'] / plat_grouped['Clicks']
    
    # Format table for visualization
    plat_formatted = plat_grouped.copy()
    plat_formatted['Spend'] = plat_formatted['Spend'].apply(lambda x: f"₹{x:,.0f}")
    plat_formatted['Revenue'] = plat_formatted['Revenue'].apply(lambda x: f"₹{x:,.0f}")
    plat_formatted['CPA'] = plat_formatted['CPA'].apply(lambda x: f"₹{x:,.0f}")
    plat_formatted['ROAS'] = plat_formatted['ROAS'].apply(lambda x: f"{x:.2f}x")
    plat_formatted['CTR'] = plat_formatted['CTR'].apply(lambda x: f"{x:.2f}%")
    plat_formatted['CPC'] = plat_formatted['CPC'].apply(lambda x: f"₹{x:.2f}")
    
    st.dataframe(plat_formatted[['Platform', 'Spend', 'Revenue', 'ROAS', 'CPA', 'CTR', 'CPC', 'Purchases']], use_container_width=True)

# --- TAB 3: PRODUCT BREAKDOWN ---
with tab_products:
    st.header("Product Performance Matrix")
    
    col_f1, col_f2 = st.columns([1, 2])
    with col_f1:
        prod_plat_filter = st.selectbox("Platform Filter", ["All Platforms (Blended)", "Meta", "Google"])
    with col_f2:
        prod_search = st.text_input("Search Product SKU")

    # Filter data
    prod_df = df.copy()
    if prod_plat_filter != "All Platforms (Blended)":
        prod_df = prod_df[prod_df['Platform'] == prod_plat_filter]
    if prod_search:
        prod_df = prod_df[prod_df['Product'].str.contains(prod_search, case=False)]

    prod_grouped = prod_df.groupby(['Product', 'Platform']).agg({
        'Spend': 'sum',
        'Revenue': 'sum',
        'Purchases': 'sum',
        'Clicks': 'sum',
        'Impressions': 'sum'
    }).reset_index()

    prod_grouped['ROAS'] = prod_grouped['Revenue'] / prod_grouped['Spend']
    prod_grouped['CPA'] = prod_grouped['Spend'] / prod_grouped['Purchases']
    prod_grouped['CTR'] = (prod_grouped['Clicks'] / prod_grouped['Impressions']) * 100
    prod_grouped['CPC'] = prod_grouped['Spend'] / prod_grouped['Clicks']

    prod_formatted = prod_grouped.copy()
    prod_formatted['Spend'] = prod_formatted['Spend'].apply(lambda x: f"₹{x:,.0f}")
    prod_formatted['Revenue'] = prod_formatted['Revenue'].apply(lambda x: f"₹{x:,.0f}")
    prod_formatted['CPA'] = prod_formatted['CPA'].apply(lambda x: f"₹{x:,.0f}" if not np.isnan(x) and not np.isinf(x) else "₹0")
    prod_formatted['ROAS'] = prod_formatted['ROAS'].apply(lambda x: f"{x:.2f}x")
    prod_formatted['CTR'] = prod_formatted['CTR'].apply(lambda x: f"{x:.2f}%")
    prod_formatted['CPC'] = prod_formatted['CPC'].apply(lambda x: f"₹{x:.2f}")

    st.dataframe(prod_formatted[['Product', 'Platform', 'Spend', 'Revenue', 'ROAS', 'CPA', 'CTR', 'CPC', 'Purchases']], use_container_width=True)

# --- TAB 4: INTELLIGENCE MODE ---
with tab_intelligence:
    st.header("⚡ Trend Analyzer & Decision Engine")
    st.markdown("Audits 3-day short-term velocity against 7-day historical benchmarks.")

    intel_rows = []
    combos = df.groupby(['Product', 'Platform']).size().reset_index()

    for idx, combo in combos.iterrows():
        prod, plat = combo['Product'], combo['Platform']
        sub_df = df[(df['Product'] == prod) & (df['Platform'] == plat)]
        
        # Split dates
        c_recent = sub_df[sub_df['Date'].isin(recent_dates)]
        c_base = sub_df[sub_df['Date'].isin(baseline_dates)]
        
        cr_spend = c_recent['Spend'].sum()
        cr_rev = c_recent['Revenue'].sum()
        cr_purch = c_recent['Purchases'].sum()
        cr_clicks = c_recent['Clicks'].sum()
        cr_impr = c_recent['Impressions'].sum()

        cb_spend = c_base['Spend'].sum()
        cb_rev = c_base['Revenue'].sum()
        cb_purch = c_base['Purchases'].sum()
        cb_clicks = c_base['Clicks'].sum()
        cb_impr = c_base['Impressions'].sum()

        # Ratios
        rec_roas = cr_rev / cr_spend if cr_spend > 0 else 0
        rec_cpa = cr_spend / cr_purch if cr_purch > 0 else 0
        rec_ctr = (cr_clicks / cr_impr * 100) if cr_impr > 0 else 0

        bas_roas = cb_rev / cb_spend if cb_spend > 0 else 0
        bas_cpa = cb_spend / cb_purch if cb_purch > 0 else 0
        bas_ctr = (cb_clicks / cb_impr * 100) if cb_impr > 0 else 0

        ctr_chg = ((rec_ctr - bas_ctr) / bas_ctr * 100) if bas_ctr > 0 else 0

        verdict = "HOLD / WATCH"
        reasoning = "Insufficient spend to compute scaling verdicts."

        if cr_spend > 200:
            if rec_roas >= target_roas and rec_cpa <= target_cpa:
                if ctr_chg >= -10:
                    verdict = "🚀 SCALE BUDGET"
                    reasoning = f"ROAS of {rec_roas:.2f}x beats target ({target_roas}x) and CPA (₹{int(rec_cpa)}) is under control with stable CTR. Scale budget 20%."
                else:
                    verdict = "🚀 SCALE (WATCH CTR)"
                    reasoning = f"Performance is solid, but CTR fell {abs(ctr_chg):.1f}% recently. Scale budget conservatively; ad set is showing early fatigue."
            elif rec_roas < 1.3 or rec_cpa > (target_cpa * 1.8) or (cr_spend > (target_cpa * 2.5) and cr_purch == 0):
                verdict = "🛑 PAUSE CREATIVE"
                if cr_purch == 0:
                    reasoning = f"Spent ₹{int(cr_spend)} (exceeded 2.5x Target CPA) in last 3 days with 0 purchases. Kill immediately."
                else:
                    reasoning = f"ROAS collapsed to {rec_roas:.2f}x and CPA spiked to ₹{int(rec_cpa)} (Target: ₹{target_cpa}). Pause budget leakage."
            elif ctr_chg < -25 and rec_cpa > bas_cpa * 1.25:
                verdict = "⚠️ CREATIVE FATIGUE"
                reasoning = f"CTR dropped by {abs(ctr_chg):.1f}% recently and CPA surged. Trim budget 30% and deploy fresh hook angles."
            else:
                verdict = "👀 HOLD / WATCH"
                reasoning = f"ROAS is {rec_roas:.2f}x (Target: {target_roas}x) and CPA is ₹{int(rec_cpa)}. Monitor closely for 48h."

        intel_rows.append({
            "Dimension": f"{prod} @ {plat}",
            "3d Spend": f"₹{int(cr_spend):,}",
            "3d ROAS": f"{rec_roas:.2f}x",
            "3d CPA": f"₹{int(rec_cpa):,}" if rec_cpa > 0 else "₹0",
            "7d ROAS": f"{bas_roas:.2f}x",
            "7d CPA": f"₹{int(bas_cpa):,}" if bas_cpa > 0 else "₹0",
            "CTR Delta": f"{ctr_chg:+.1f}%",
            "Verdict": verdict,
            "Reasoning": reasoning
        })

    st.dataframe(pd.DataFrame(intel_rows), use_container_width=True)

# --- TAB 5: COMPETITOR SPY ---
with tab_competitors:
    st.header("Competitor Creative Swipes")
    
    col_sw1, col_sw2 = st.columns([1, 2])
    
    with col_sw1:
        st.subheader("Log Competitor Angle")
        c_name = st.text_input("Competitor Name")
        c_plat = st.selectbox("Platform", ["Meta", "Google", "TikTok", "YouTube"])
        c_hook = st.selectbox("Hook Format", [
            "Us vs Them (Comparison Table)",
            "Three Reasons Why (Feature Callout)",
            "Scientific Diagnostic / Doctor Reaction",
            "UGC Extreme Pain Point Demonstration",
            "Price/Discount Flash Promotion"
        ])
        c_link = st.text_input("Library Ad URL")
        c_copy = st.text_area("Ad Copy Angle", placeholder="Copy pasted hooks here...")
        c_insights = st.text_area("Defensive Counter-Angle", placeholder="How our brand counters this messaging hook...")
        
        if st.button("➕ Log Hook"):
            if c_name and c_link and c_copy:
                st.session_state.competitor_ads.insert(0, {
                    "Name": c_name,
                    "Platform": c_plat,
                    "Hook": c_hook,
                    "Link": c_link,
                    "Copy": c_copy,
                    "Insights": c_insights
                })
                st.success("Competitor hook added!")
                st.rerun()
            else:
                st.warning("Please fill all required inputs (Name, URL, Hook copy).")
                
    with col_sw2:
        st.subheader("Spy swipe library")
        
        for idx, item in enumerate(st.session_state.competitor_ads):
            with st.container():
                # Render competitor cards
                st.markdown(f"""
                <div style="background-color:rgba(15, 23, 42, 0.4); border:1px solid rgba(255,255,255,0.05); padding:15px; border-radius:10px; margin-bottom:15px;">
                    <h5 style="margin:0; font-family:'Space Grotesk'; color:#fff;">🕵️ {item['Name']} <span style="font-size:10px; background-color:#6366f1; padding:2px 6px; border-radius:10px;">{item['Platform']}</span></h5>
                    <div style="font-size:12px; margin-top:8px; color:#9ca3af;">
                        <span style="color:#f59e0b; font-weight:600;">Format:</span> {item['Hook']}<br>
                        <span style="color:#cbd5e1;"><b>Hook Copy:</b></span> <i>"{item['Copy']}"</i><br>
                        <span style="color:#10b981;"><b>Defensive Counter:</b></span> {item['Insights']}<br>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                if st.button("Delete Hook", key=f"del_{idx}"):
                    st.session_state.competitor_ads.pop(idx)
                    st.rerun()

# --- TAB 6: GROWTH AGENT ---
with tab_agent:
    st.header("Interactive Growth Director Chat")
    
    # Message container
    chat_container = st.container()
    
    with chat_container:
        for msg in st.session_state.chat_history:
            if msg["role"] == "user":
                st.chat_message("user").write(msg["content"])
            else:
                st.chat_message("assistant").write(msg["content"])

    user_query = st.chat_input("Type your marketing query here...")
    
    if user_query:
        # Save user message
        st.session_state.chat_history.append({"role": "user", "content": user_query})
        st.rerun()

    # Process responses on rerun after new input
    if len(st.session_state.chat_history) > 0 and st.session_state.chat_history[-1]["role"] == "user":
        latest_user_msg = st.session_state.chat_history[-1]["content"]
        
        with st.spinner("AI Director is auditing metrics..."):
            if api_key and GENAI_AVAILABLE:
                try:
                    # Prepare context and compile prompt
                    clean_intel = []
                    for idx, combo in combos.iterrows():
                        prod, plat = combo['Product'], combo['Platform']
                        sub_df = df[(df['Product'] == prod) & (df['Platform'] == plat)]
                        c_recent = sub_df[sub_df['Date'].isin(recent_dates)]
                        cr_spend = c_recent['Spend'].sum()
                        cr_rev = c_recent['Revenue'].sum()
                        cr_purch = c_recent['Purchases'].sum()
                        rec_roas = cr_rev / cr_spend if cr_spend > 0 else 0
                        rec_cpa = cr_spend / cr_purch if cr_purch > 0 else 0
                        
                        clean_intel.append({
                            "Product": prod,
                            "Platform": plat,
                            "Spend": cr_spend,
                            "ROAS": rec_roas,
                            "CPA": rec_cpa
                        })
                    
                    sys_prompt = f"""
                    Act as Marin, an elite D2C growth marketing and retention director. You speak with extreme clarity, focus on profit margins, and deliver prescriptive action plans without fluff.
                    
                    You have analyzed the account metrics:
                    Blended Spend (3d): ₹{r_spend}
                    Blended Revenue (3d): ₹{r_rev}
                    Blended ROAS (3d): {recent_roas:.2f}x
                    Blended CPA (3d): ₹{recent_cpa:.0f}
                    
                    Product-Level 3d reports:
                    {json.dumps(clean_intel, indent=2)}
                    
                    Competitor Library hooks:
                    {json.dumps(st.session_state.competitor_ads, indent=2)}
                    
                    Target Benchmarks: Target ROAS = {target_roas}x, Target CPA = ₹{target_cpa}
                    
                    Answer the user's growth query. Link your advice directly to the metrics above, citing specific product ROAS and CPA values. Highlight action points in bold.
                    """
                    
                    client = genai.Client(api_key=api_key)
                    response = client.models.generate_content(
                        model="gemini-2.5-flash",
                        contents=[sys_prompt, f"User query: {latest_user_msg}"]
                    )
                    
                    st.session_state.chat_history.append({"role": "agent", "content": response.text})
                    st.rerun()
                except Exception as e:
                    st.error(f"API Error: {str(e)}. Falling back to local heuristic rules.")
                    time.sleep(1)
            
            # Local Heuristic Response Fallback
            normalized = latest_user_msg.lower()
            resp = "I have run a local mathematical check. "
            
            if "scale" in normalized or "winner" in normalized:
                # Find products with good metrics
                scalers = []
                for idx, combo in combos.iterrows():
                    prod, plat = combo['Product'], combo['Platform']
                    sub_df = df[(df['Product'] == prod) & (df['Platform'] == plat)]
                    c_recent = sub_df[sub_df['Date'].isin(recent_dates)]
                    cr_spend = c_recent['Spend'].sum()
                    cr_rev = c_recent['Revenue'].sum()
                    cr_purch = c_recent['Purchases'].sum()
                    rec_roas = cr_rev / cr_spend if cr_spend > 0 else 0
                    rec_cpa = cr_spend / cr_purch if cr_purch > 0 else 0
                    
                    if cr_spend > 200 and rec_roas >= target_roas and rec_cpa <= target_cpa:
                        scalers.append(f"• **{prod} ({plat} Ads)**: ROAS is **{rec_roas:.2f}x** (above target {target_roas}x), CPA is **₹{int(rec_cpa)}**.")
                
                if scalers:
                    resp += "We recommend scaling budgets on these winners by 20%:\n\n" + "\n".join(scalers)
                else:
                    resp += "No product-platform combination currently meets target metrics for scaling."
            elif "pause" in normalized or "fatigue" in normalized or "waste" in normalized:
                pausers = []
                for idx, combo in combos.iterrows():
                    prod, plat = combo['Product'], combo['Platform']
                    sub_df = df[(df['Product'] == prod) & (df['Platform'] == plat)]
                    c_recent = sub_df[sub_df['Date'].isin(recent_dates)]
                    cr_spend = c_recent['Spend'].sum()
                    cr_rev = c_recent['Revenue'].sum()
                    cr_purch = c_recent['Purchases'].sum()
                    rec_roas = cr_rev / cr_spend if cr_spend > 0 else 0
                    rec_cpa = cr_spend / cr_purch if cr_purch > 0 else 0
                    
                    if cr_spend > 200 and (rec_roas < 1.3 or rec_cpa > (target_cpa * 1.5)):
                        pausers.append(f"• **{prod} ({plat} Ads)**: ROAS collapsed to **{rec_roas:.2f}x**, CPA surged to **₹{int(rec_cpa)}**.")
                
                if pausers:
                    resp += "We recommend pausing or reducing budget on these underperformers to stop leakage:\n\n" + "\n".join(pausers)
                else:
                    resp += "Excellent news! No campaigns are currently exceeding our pause thresholds."
            else:
                resp += f"""Here is your blended scorecard context:\n
                - **Blended ROAS**: {recent_roas:.2f}x (Target: {target_roas}x)
                - **Blended CPA**: ₹{int(recent_cpa)} (Target: ₹{target_cpa})
                - **Blended Spend**: ₹{int(r_spend):,}
                \nAsk me to review what to scale, what to pause, or design copy variations! (Enter your API Key in the sidebar to activate reasoning insights)."""
                
            st.session_state.chat_history.append({"role": "agent", "content": resp})
            st.rerun()
