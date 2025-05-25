import streamlit as st
import pandas as pd
import requests
import time
from urllib.parse import urlparse
from datetime import datetime
from io import BytesIO
import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe

# --- CONFIG ---
API_KEY = "your_serpapi_key_here"
TARGET_DOMAIN = "kollegeapply.com"
SPREADSHEET_ID = "1tzdEOUYqFRZAizMkKHFyYzoYHv4t_XOt_RauNlMt9QE"

COMPETITORS = {
    "shiksha": "shiksha.com",
    "collegedunia": "collegedunia.com",
    "collegedekho": "collegedekho.com"
}

OFFICIAL_PATTERNS = ["wikipedia.org", "wikimedia.org", "britannica.com", ".europa.eu"]

AMBIGUOUS_QUERIES = {"cat": "CAT exam", "gmat": "GMAT exam", "gre": "GRE exam"}


def is_official_site(url):
    return any(p in url.lower() for p in OFFICIAL_PATTERNS)

def extract_domain(url):
    try:
        return urlparse(url).netloc.lower().replace("www.", "")
    except:
        return ""

def domain_in_url(url, domain):
    try:
        return domain in urlparse(url).netloc.lower()
    except:
        return False

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
                    row[f"rank_{i}_name"] = filtered[i - 1].get("title", "")
                    row[f"rank_{i}_url"] = filtered[i - 1].get("link", "")
                else:
                    row[f"rank_{i}_name"] = ""
                    row[f"rank_{i}_url"] = ""

            for name, dom in COMPETITORS.items():
                rnk, url = get_ranking(filtered, dom)
                row[f"{name}_rank"] = rnk
                row[f"{name}_url"] = url

            paa_results = data.get("related_questions", [])
            row["paa_exists"] = "Yes" if paa_results else "No"
            row["paa_kollegeapply"] = "No"
            for item in paa_results:
                source = item.get("source", {})
                if source and domain_in_url(source.get("link", ""), TARGET_DOMAIN):
                    row["paa_kollegeapply"] = "Yes"
                    break

            results.append(row)

        except Exception as e:
            st.error(f"❌ Error for keyword '{kw}': {e}")
        time.sleep(2)

    return pd.DataFrame(results)


# --- UI ---
st.set_page_config(page_title="SERP Rank + PAA Tracker", layout="wide")
st.title("🔍 Keyword Rank Tracker + PAA Checker")

uploaded_file = st.file_uploader("Upload Excel with 'KW' column", type=["xlsx"])

if uploaded_file:
    try:
        df_kw = pd.read_excel(uploaded_file, usecols=["KW"])
        st.success("✅ File uploaded successfully")

        with st.spinner("Processing..."):
            df_result = process_keywords(df_kw)

            # Authenticate with Google Sheets
            creds = Credentials.from_service_account_file(
                "service_account.json",
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            gc = gspread.authorize(creds)
            sh = gc.open_by_key(SPREADSHEET_ID)

            sheet_name = datetime.now().strftime("Rank_%Y%m%d_%H%M")
            worksheet = sh.add_worksheet(title=sheet_name, rows="100", cols="30")
            set_with_dataframe(worksheet, df_result)

        st.success(f"✅ Data written to Google Sheet tab: {sheet_name}")
        st.dataframe(df_result.head(10))

        # Optional: download Excel
        towrite = BytesIO()
        df_result.to_excel(towrite, index=False, engine='openpyxl')
        towrite.seek(0)
        st.download_button("📥 Download Excel", towrite, "keyword_rank_output.xlsx")

    except Exception as e:
        st.error(f"⚠️ {e}")
