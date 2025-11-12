"""
update_research_data.py
-------------------------------------------------------
- Reads Excel of faculty (Faculty Name, Google Scholar Profile URL)
- Fetches latest 2–3 publications per faculty via scholarly
- Produces faculty_publications.json (array of objects)
- No server upload — file saved locally in /output
- 100% 404-safe: all links go to faculty's main Google Scholar profile
-------------------------------------------------------
"""

import os
import json
import time
import pandas as pd
from scholarly import scholarly  # pip install scholarly
from urllib.parse import urlparse, parse_qs

# -------------------------
# CONFIGURATION — EDIT THESE
# -------------------------
EXCEL_FILENAME = "Google Scholar and ORCID Data.xlsx"  # Excel file name
SEARCH_SUBFOLDERS = True  # search recursively for Excel
OUTPUT_DIR = "output"
PAPERS_PER_FACULTY = 5  # change to 2 if you only want top 2 papers

# -------------------------
# FIND EXCEL FILE
# -------------------------
def find_excel_file(filename, search_subfolders=True):
    """Find Excel file automatically in the current project folder."""
    for root, dirs, files in os.walk(os.getcwd()):
        if filename in files:
            return os.path.join(root, filename)
        if not search_subfolders:
            break
    raise FileNotFoundError(f"❌ Excel file '{filename}' not found in this project folder.")

EXCEL_PATH = find_excel_file(EXCEL_FILENAME, SEARCH_SUBFOLDERS)
print(f"✅ Excel file detected at: {EXCEL_PATH}")

# Prepare output folder
os.makedirs(OUTPUT_DIR, exist_ok=True)
OUTPUT_JSON = os.path.join(OUTPUT_DIR, "faculty_publications.json")

# -------------------------
# HELPER FUNCTIONS
# -------------------------
def extract_scholar_userid(url):
    """Extracts the Google Scholar user ID from a profile URL."""
    try:
        qs = parse_qs(urlparse(url).query)
        if "user" in qs:
            return qs["user"][0]
        for part in url.split("?"):
            if "user=" in part:
                return part.split("user=")[1].split("&")[0]
    except Exception:
        pass
    return None


def fetch_author_pubs(user_id, scholar_url, n=3):
    """Fetches n latest publications for a given Google Scholar user (404-safe version)."""
    try:
        author = scholarly.search_author_id(user_id)
        author = scholarly.fill(author, sections=["publications"])
        pubs = author.get("publications", [])
        structured = []

        for pub in pubs:
            bib = pub.get("bib", {})
            year = bib.get("pub_year") or bib.get("year") or 0
            title = bib.get("title", "N/A")
            venue = bib.get("venue") or bib.get("journal") or bib.get("publisher") or ""
            citations = pub.get("num_citations", 0)

            # ✅ Safe: link always points to the faculty's Scholar profile (no 404)
            link = f"https://scholar.google.com/citations?user={user_id}" if user_id else scholar_url

            structured.append({
                "Title": title,
                "Journal": venue,
                "Year": int(year) if str(year).isdigit() else year,
                "Citations": int(citations) if citations else 0,
                "Link": link
            })

        # Sort by year (descending) and limit to n results
        structured_sorted = sorted(structured, key=lambda x: (x.get("Year") or 0), reverse=True)
        return structured_sorted[:n]

    except Exception as e:
        print(f"⚠️ Error fetching publications for user {user_id}: {e}")
        return []


# -------------------------
# MAIN FUNCTION
# -------------------------
def main():
    print("\n🔍 Reading Excel file...")
    df = pd.read_excel(EXCEL_PATH)
    print(f"✅ Excel loaded successfully with {len(df)} rows.\n")

    expected_cols = ["Faculty Name", "Google Scholar Profile URL"]
    if not all(col in df.columns for col in expected_cols):
        print(f"⚠️ Missing expected columns. Found: {list(df.columns)}")
        print(f"Please ensure your Excel has columns: {expected_cols}")
        return

    results = []
    print(f"🔎 Found {len(df)} faculty records. Starting publication fetch...\n")

    for _, row in df.iterrows():
        name = row["Faculty Name"]
        scholar_url = row["Google Scholar Profile URL"]

        print(f"👉 Fetching for: {name}")
        user_id = extract_scholar_userid(scholar_url)

        if not user_id:
            print(f"⚠️ Skipping {name}: Invalid Google Scholar URL.\n")
            continue

        pubs = fetch_author_pubs(user_id, scholar_url, n=PAPERS_PER_FACULTY)
        print(f"✅ {len(pubs)} publications fetched for {name}.\n")

        for p in pubs:
            results.append({
                "Faculty Name": name,
                "Title": p["Title"],
                "Journal": p["Journal"],
                "Year": p["Year"],
                "Citations": p["Citations"],
                "Link": p["Link"]
            })

        time.sleep(2)  # polite pause to avoid hitting Scholar limits

    # Save results locally
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    print(f"\n✅ JSON file created successfully!")
    print(f"📄 File location: {os.path.abspath(OUTPUT_JSON)}\n")
    print("➡️ Upload this file manually to your GitHub repository or website.")
    print("   Example: https://<yourusername>.github.io/faculty-publications/faculty_publications.json\n")


# -------------------------
# RUN SCRIPT
# -------------------------
if __name__ == "__main__":
    main()

