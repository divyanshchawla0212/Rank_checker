# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import requests
import time
from urllib.parse import urlparse
from io import BytesIO

API_KEY = "99ecb6787f81d593d52c5022c099144087de6c84e4cbf9c071118e9af0810da2"
TARGET_DOMAIN = "kollegeapply.com"

COMPETITORS = {
    "shiksha": "shiksha.com",
    "collegedunia": "collegedunia.com",
    "collegedekho": "collegedekho.com"
}

OFFICIAL_PATTERNS = [
    "wikipedia.org", "wikimedia.org", "britannica.com", ".europa.eu"
]

AMBIGUOUS_QUERIES = {
    "cat": "CAT exam",
    "gmat": "GMAT exam",
    "gre": "GRE exam",
}

def is_official_site(url):
    u = url.lower()
    return any(p in u for p in OFFICIAL_PATTERNS)

def extract_domain(url):
    try:
        return urlparse(url).netloc.lower().replace("www.", "")
    except:
        return ""

def get_ranking(organic_results, target_domain):
    for idx, r in enumerate(organic_results, start=1):
        link = r.get("link", "")
        domain = extract_domain(link)
        if target_domain in domain:
            return idx, link
    return "Not in Top 100", ""

def process_keywords(df_kw):
    results = []

    for kw in df_kw["KW"]:
        query = AMBIGUOUS_QUERIES.get(kw.strip().lower(), kw)

        params = {
            "engine": "google",
            "q": query,
            "api_key": API_KEY,
            "num": 100,
            "hl": "en",
            "gl": "in",
            "google_domain": "google.co.in",
            "device": "desktop"
        }

        try:
            resp = requests.get("https://serpapi.com/search", params=params)
            resp.raise_for_status()
            data = resp.json()
            organic = data.get("organic_results", [])
            filtered = [r for r in organic if not is_official_site(r.get("link", ""))]

            row = {"keyword": kw}
            kr, ku = get_ranking(filtered, TARGET_DOMAIN)
            row["kollegeapply_rank"] = kr
            row["kollegeapply_url"] = ku

            for i in range(1, 4):
                if len(filtered) >= i:
                    row[f"rank_{i}_name"] = filtered[i-1].get("title", "")
                    row[f"rank_{i}_url"] = filtered[i-1].get("link", "")
                else:
                    row[f"rank_{i}_name"] = ""
                    row[f"rank_{i}_url"] = ""

            for name, dom in COMPETITORS.items():
                rnk, url = get_ranking(filtered, dom)
                row[f"{name}_rank"] = rnk
                row[f"{name}_url"] = url

            results.append(row)

        except requests.exceptions.RequestException as e:
            st.error(f"❌ Error processing keyword '{kw}': {e}")
        
        time.sleep(2)  # Delay to avoid hitting SerpAPI rate limits

    return pd.DataFrame(results)

# --- Streamlit App ---
st.title("📈 Google Keyword Rank Checker")

uploaded_file = st.file_uploader("Upload Keyword Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df_kw = pd.read_excel(uploaded_file, usecols=["KW"])
        st.success("✅ File uploaded. Starting processing...")

        with st.spinner("🔄 Fetching data from Google..."):
            df_results = process_keywords(df_kw)

        st.subheader("🔍 Preview of Results")
        st.dataframe(df_results.head(10))

        # Download
        towrite = BytesIO()
        df_results.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button("📥 Download Excel", towrite, "keyword_ranks.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    except Exception as e:
        st.error(f"⚠️ Error reading file: {e}")
