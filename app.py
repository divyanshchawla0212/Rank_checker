import streamlit as st
import pandas as pd
import requests
import time
from urllib.parse import urlparse
from io import BytesIO
from datetime import datetime

# --- CONFIGURATION ---
API_KEY = "8ae51a89b5c6bacd1e3b1783c1f7cacae0eed213d80afebf43a11018d27ee885"
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
    "gre": "GRE exam"
}

# --- HELPER FUNCTIONS ---
def is_official_site(url):
    return any(p in url.lower() for p in OFFICIAL_PATTERNS)

def extract_domain(url):
    try:
        return urlparse(url).netloc.lower().replace("www.", "")
    except:
        return ""

def domain_in_url(url, domain):
    try:
        return domain in extract_domain(url)
    except:
        return False

def get_ranking(organic_results, target_domain):
    for idx, r in enumerate(organic_results, start=1):
        link = r.get("link", "")
        domain = extract_domain(link)
        if target_domain in domain:
            return idx, link
    return "Not in Top 100", ""

# --- MAIN PROCESSING FUNCTION ---
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
                    row[f"rank_{i}_name"] = filtered[i - 1].get("title", "")
                    row[f"rank_{i}_url"] = filtered[i - 1].get("link", "")
                else:
                    row[f"rank_{i}_name"] = ""
                    row[f"rank_{i}_url"] = ""

            for name, dom in COMPETITORS.items():
                rnk, url = get_ranking(filtered, dom)
                row[f"{name}_rank"] = rnk
                row[f"{name}_url"] = url

            # --- Featured Snippet + PAA Detection ---
            row["paa_exists"] = "No"
            row["paa_kollegeapply"] = "No"
            checked_links = []

            # 1. Featured snippet / answer box
            featured = data.get("answer_box", {}) or data.get("featured_snippet", {})
            fs_link = featured.get("link", "")
            if fs_link:
                checked_links.append(fs_link)
                if domain_in_url(fs_link, TARGET_DOMAIN):
                    row["paa_exists"] = "Yes"
                    row["paa_kollegeapply"] = "Yes"

            # 2. PAA (People Also Ask)
            paa_results = data.get("related_questions", [])
            if paa_results:
                row["paa_exists"] = "Yes"
                for item in paa_results:
                    link = ""
                    if "source" in item and "link" in item["source"]:
                        link = item["source"]["link"]
                    elif "answer" in item and "source" in item["answer"] and "link" in item["answer"]["source"]:
                        link = item["answer"]["source"]["link"]
                    if link:
                        checked_links.append(link)
                        if domain_in_url(link, TARGET_DOMAIN):
                            row["paa_kollegeapply"] = "Yes"
                            break

            # 3. Add debug info of all links checked
            row["paa_links_checked"] = "; ".join(checked_links)

            results.append(row)

        except Exception as e:
            st.error(f"❌ Error processing keyword '{kw}': {e}")

        time.sleep(2)

    return pd.DataFrame(results)

# --- STREAMLIT UI ---
st.set_page_config(page_title="Keyword Rank + PAA Checker", layout="wide")
st.title("🔍 Keyword SERP Rank Checker + PAA/Featured Snippet Detection")

uploaded_file = st.file_uploader("📤 Upload Excel file with a 'KW' column", type=["xlsx"])

if uploaded_file:
    try:
        df_kw = pd.read_excel(uploaded_file, usecols=["KW"])
        st.success("✅ File uploaded. Processing...")

        with st.spinner("🔄 Querying Google SERPs via SerpAPI..."):
            df_results = process_keywords(df_kw)

        st.subheader("📊 Result Preview")
        st.dataframe(df_results.head(10))

        # Download
        towrite = BytesIO()
        df_results.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"keyword_rank_output_{timestamp}.xlsx"
        st.download_button("📥 Download Excel", towrite, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"⚠️ Failed to process file: {e}")
