#!/usr/bin/env python3
"""
Daily job refresh for Job_Tracker_Mumbai_Java.xlsx

Priority order in the sheet:
  1 Mumbai / Navi Mumbai / Thane / Remote
  2 Verify location (search result didn't show a city)
  3 Other cities
Within each group: High fit first, then newest posting first.

What a run does
  1. Runs Firecrawl searches (LinkedIn, Naukri, web), limited to the past week.
  2. Drops WITCH companies (Wipro, Infosys, TCS, Cognizant, HCL), junior/too-senior
     roles, and anything it can see is older than 7 days.
  3. Adds only NEW jobs (deduplicated by link), then re-sorts the whole Jobs sheet.
     Your Status / Date Applied / Notes / Fit edits are preserved.

Setup
  pip install requests openpyxl
  export FIRECRAWL_API_KEY="fc-..."
  python job_refresh.py
"""
import os
import re
import sys
from datetime import date, datetime, timedelta

import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side

API_KEY = os.environ.get("FIRECRAWL_API_KEY", "")
TRACKER_PATH = os.environ.get("TRACKER_PATH", "Job_Tracker_Mumbai_Java.xlsx")
SEARCH_URL = "https://api.firecrawl.dev/v2/search"
MAX_AGE_DAYS = 7
MAX_MIN_EXPERIENCE = 7   # skip roles whose minimum experience is above this
MIN_MAX_EXPERIENCE = 3   # skip roles whose maximum experience is below this

# (query, location bias). Group 1 = Mumbai + remote, group 2 = other cities.
QUERIES = [
    ("site:linkedin.com/jobs Java Spring Boot backend engineer Mumbai", "Mumbai, India"),
    ("site:linkedin.com/jobs Java Kafka microservices engineer Mumbai", "Mumbai, India"),
    ("site:linkedin.com/jobs Software Engineer II OR III OR SDE 2 Java Mumbai", "Mumbai, India"),
    ("site:naukri.com Java Spring Boot developer Mumbai 4 to 8 years", "Mumbai, India"),
    ("site:naukri.com Java backend Kafka microservices Mumbai", "Mumbai, India"),
    ("SDE 2 backend Java Spring Boot Kafka Mumbai fintech product company", "Mumbai, India"),
    ("site:linkedin.com/jobs Java Spring Boot backend engineer remote India", "India"),
    ("site:naukri.com Java Spring Boot developer work from home remote 4 to 8 years", "India"),
    ("site:linkedin.com/jobs Java Spring Boot Kafka backend engineer Pune OR Bengaluru OR Hyderabad", "India"),
    ("site:naukri.com Java Spring Boot Kafka 4 to 8 years Pune Bengaluru Hyderabad", "India"),
]

WITCH = re.compile(
    r"\b(wipro|infosys|tata consultancy|tcs|cognizant|hcl|hcltech|hcl technologies)\b"
)
NOT_RELEVANT_TITLE = re.compile(
    r"\b(intern|internship|fresher|freshers|trainee|junior|principal|architect|"
    r"director|manager|vp|vice president)\b"
)
MUST_MATCH = re.compile(r"\b(java|spring|backend|back end|sde|software engineer)\b")
MUMBAI = re.compile(r"\b(mumbai|navi mumbai|thane|bombay)\b")
REMOTE = re.compile(r"\b(remote|work from home|wfh)\b")
CITIES = {
    "bengaluru": "Bengaluru", "bangalore": "Bengaluru", "pune": "Pune",
    "hyderabad": "Hyderabad", "chennai": "Chennai", "gurgaon": "Gurugram",
    "gurugram": "Gurugram", "noida": "Noida", "delhi": "Delhi", "kolkata": "Kolkata",
    "ahmedabad": "Ahmedabad", "kochi": "Kochi", "indore": "Indore", "jaipur": "Jaipur",
    "coimbatore": "Coimbatore", "nagpur": "Nagpur", "vadodara": "Vadodara",
    "chandigarh": "Chandigarh", "trivandrum": "Trivandrum", "bhubaneswar": "Bhubaneswar",
}

# ---------- sheet layout (shared with the tracker file) ----------
HEADERS = ["#", "Priority", "Company", "Role", "Location", "Experience", "Source",
           "Approx. Posted", "Age (days)", "Fit", "Status", "Date Applied", "Link",
           "Notes", "Date Found"]
(COL_NUM, COL_PRI, COL_COMP, COL_ROLE, COL_LOC, COL_EXP, COL_SRC, COL_POSTED, COL_AGE,
 COL_FIT, COL_STATUS, COL_APPLIED, COL_LINK, COL_NOTE, COL_FOUND) = range(1, 16)
LAST_COL = COL_FOUND
FONT = "Arial"
THIN = Side(style="thin", color="BFBFBF")
FIT_RANK = {"High": 0, "Medium": 1, "Low": 2}


# ---------- helpers ----------
def norm(text):
    return re.sub(r"[-_/]+", " ", (text or "").lower())


def clean_url(url):
    return url.split("?")[0].rstrip("/")


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    return v if isinstance(v, date) else None


def source_of(url):
    if "linkedin.com" in url:
        return "LinkedIn"
    if "naukri.com" in url:
        return "Naukri"
    return "Other"


def posted_date(title, desc, url, today):
    """Best-effort posting date; None if unknown."""
    text = f"{title} {desc}".lower()
    m = re.search(r"(\d+)\s*(hour|hr|minute|min|day|week)s?\s+ago", text)
    if m:
        n, unit = int(m.group(1)), m.group(2)
        if unit in ("hour", "hr", "minute", "min"):
            return today
        return today - timedelta(days=n if unit == "day" else 7 * n)
    # Naukri job ids end with DDMMYY + digits, e.g. ...-121026002418 -> 12 Oct 2026
    m = re.search(r"-(\d{2})(\d{2})(\d{2})\d{4,7}$", url.split("?")[0])
    if "naukri.com" in url and m:
        try:
            return date(2000 + int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None


def parse_experience(title, desc, url):
    """Returns (lo, hi) years, (lo, None) for 'N+', or None."""
    raw = f"{title} {desc} {url.replace('-', ' ')}".lower()
    m = re.search(r"\b(\d{1,2})\s*(?:to|-|–)\s*(\d{1,2})\s*\+?\s*(?:yrs|years|yr)\b", raw)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r"\b(\d{1,2})\s*\+\s*(?:yrs|years|yr)\b", raw)
    if m:
        return int(m.group(1)), None
    return None


def exp_label(exp):
    if not exp:
        return "Not stated"
    lo, hi = exp
    return f"{lo}-{hi} yrs" if hi else f"{lo}+ yrs"


def classify_location(title, url, desc):
    """Mumbai / Remote / city name(s) / Verify. Title+URL are trusted before the snippet."""
    for text in (norm(f"{title} {url}"), norm(desc)):
        if MUMBAI.search(text):
            return "Mumbai"
        if REMOTE.search(text):
            return "Remote"
        found = sorted({v for k, v in CITIES.items() if re.search(rf"\b{k}\b", text)})
        if found:
            return " / ".join(found)
    return "Verify"


def tier(location):
    """1 = Mumbai/Remote, 2 = unknown, 3 = other city. Mirrors the Priority formula."""
    loc = (location or "").lower()
    if any(k in loc for k in ("mumbai", "thane", "remote")):
        return 1
    if loc in ("", "verify"):
        return 2
    return 3


def guess_company_role(title):
    """Very rough; the full title still ends up in Role. Fix company by hand if needed."""
    t = re.sub(r"\s*[-|]\s*LinkedIn( Jobs)?\s*$", "", title).strip()
    m = re.match(r"(.+?)\s+hiring\s+(.+?)\s+in\s+", t, re.I)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m = re.match(r"(.+?)\s+at\s+(.+?)(?:\s+[-—–]\s+.*)?$", t, re.I)
    if m:
        return m.group(2).strip(), m.group(1).strip()
    m = re.match(r"(.+?)\s+[-–]\s+(.+?)\s+-\s+\d+", t)  # Naukri style
    if m:
        return m.group(2).strip(), m.group(1).strip()
    return "See link", t


def fit_of(text):
    t = text.lower()
    return "High" if ("kafka" in t and "spring" in t) else "Medium"


# ---------- search ----------
def run_search(query, location):
    resp = requests.post(
        SEARCH_URL,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        json={"query": query, "limit": 20, "sources": ["web"], "tbs": "qdr:w",
              "location": location},
        timeout=90,
    )
    resp.raise_for_status()
    data = resp.json().get("data", {})
    return data.get("web", []) if isinstance(data, dict) else data


def process_results(results, seen_links, today):
    """Filter raw search results into new job records. Pure function (easy to test)."""
    out = []
    for r in results:
        url = clean_url(r.get("url", ""))
        title = r.get("title", "") or ""
        desc = r.get("description", "") or r.get("snippet", "") or ""
        blob = norm(f"{title} {desc} {url}")

        if not url or url in seen_links:
            continue
        if source_of(url) != "Other" and "/jobs/view/" not in url and "job-listings" not in url:
            continue  # category / company pages
        if WITCH.search(blob) or NOT_RELEVANT_TITLE.search(norm(title)):
            continue
        if not MUST_MATCH.search(blob):
            continue

        posted = posted_date(title, desc, url, today)
        if posted and (today - posted).days > MAX_AGE_DAYS:
            continue
        if source_of(url) == "Other" and posted is None:
            continue

        exp = parse_experience(title, desc, url)
        if exp:
            lo, hi = exp
            if lo > MAX_MIN_EXPERIENCE or (hi is not None and hi < MIN_MAX_EXPERIENCE):
                continue

        company, role = guess_company_role(title)
        out.append(dict(
            company=company, role=role, location=classify_location(title, url, desc),
            exp=exp_label(exp), source=source_of(url), posted=posted, fit=fit_of(blob),
            status="Not applied", applied=None, link=url, found=today,
            note="Auto-added; check company/location" + ("" if posted else "; posting date unknown"),
        ))
        seen_links.add(url)
    return out


# ---------- Excel ----------
def read_records(ws):
    recs = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, COL_COMP).value in (None, ""):
            continue
        link_cell = ws.cell(r, COL_LINK)
        recs.append(dict(
            company=ws.cell(r, COL_COMP).value, role=ws.cell(r, COL_ROLE).value,
            location=ws.cell(r, COL_LOC).value or "Verify",
            exp=ws.cell(r, COL_EXP).value, source=ws.cell(r, COL_SRC).value,
            posted=to_date(ws.cell(r, COL_POSTED).value), fit=ws.cell(r, COL_FIT).value,
            status=ws.cell(r, COL_STATUS).value or "Not applied",
            applied=to_date(ws.cell(r, COL_APPLIED).value),
            link=clean_url(link_cell.hyperlink.target) if link_cell.hyperlink else "",
            note=ws.cell(r, COL_NOTE).value, found=to_date(ws.cell(r, COL_FOUND).value),
        ))
    return recs


def sort_key(rec):
    posted = rec["posted"] or date.min
    return (tier(rec["location"]), FIT_RANK.get(rec["fit"], 1), -posted.toordinal())


def style_row(ws, r):
    border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    wrap_cols = (COL_COMP, COL_ROLE, COL_LOC, COL_NOTE)
    center_cols = (COL_NUM, COL_PRI, COL_EXP, COL_SRC, COL_POSTED, COL_AGE, COL_FIT,
                   COL_STATUS, COL_APPLIED, COL_LINK, COL_FOUND)
    for c in range(1, LAST_COL + 1):
        cell = ws.cell(r, c)
        cell.font = Font(name=FONT, size=10)
        cell.border = border
        cell.alignment = Alignment(
            vertical="top", wrap_text=c in wrap_cols,
            horizontal="center" if c in center_cols else None)
    for c in (COL_POSTED, COL_APPLIED, COL_FOUND):
        ws.cell(r, c).number_format = "dd-mmm-yy"
    ws.cell(r, COL_AGE).number_format = "0"
    link = ws.cell(r, COL_LINK)
    link.font = Font(name=FONT, size=10, color="0563C1", underline="single")


def write_records(ws, recs):
    recs = sorted(recs, key=sort_key)
    for r in range(2, max(ws.max_row, len(recs) + 1) + 1):
        for c in range(1, LAST_COL + 1):
            cell = ws.cell(r, c)
            cell.hyperlink = None
            cell.value = None
    for i, rec in enumerate(recs, start=2):
        style_row(ws, i)
        vals = {
            COL_NUM: i - 1,
            COL_PRI: (f'=IF(OR(ISNUMBER(SEARCH("mumbai",E{i})),ISNUMBER(SEARCH("thane",E{i})),'
                      f'ISNUMBER(SEARCH("remote",E{i}))),"1 Mumbai/Remote",'
                      f'IF(OR(E{i}="",E{i}="Verify"),"2 Verify location","3 Other city"))'),
            COL_COMP: rec["company"], COL_ROLE: rec["role"], COL_LOC: rec["location"],
            COL_EXP: rec["exp"] or "Not stated", COL_SRC: rec["source"],
            COL_POSTED: rec["posted"], COL_AGE: f'=IF(H{i}="","",TODAY()-H{i})',
            COL_FIT: rec["fit"], COL_STATUS: rec["status"], COL_APPLIED: rec["applied"],
            COL_NOTE: rec["note"], COL_FOUND: rec["found"],
        }
        for c, v in vals.items():
            ws.cell(i, c).value = v
        if rec["link"]:
            link = ws.cell(i, COL_LINK)
            link.hyperlink = rec["link"]
            link.value = "Open"
    ws.auto_filter.ref = f"A1:{ws.cell(1, LAST_COL).column_letter}{len(recs) + 1}"
    return len(recs)


def main():
    if not API_KEY:
        sys.exit("Set FIRECRAWL_API_KEY first.")
    today = date.today()
    wb = load_workbook(TRACKER_PATH)
    ws = wb["Jobs"]
    if ws.cell(1, COL_PRI).value != "Priority":
        sys.exit("Tracker layout is outdated - use the latest Job_Tracker_Mumbai_Java.xlsx.")
    recs = read_records(ws)
    seen = {r["link"] for r in recs if r["link"]}
    new = []
    for query, location in QUERIES:
        try:
            new += process_results(run_search(query, location), seen, today)
        except Exception as e:  # keep going if one query fails
            print(f"[warn] query failed: {query!r}: {e}")
    write_records(ws, recs + new)
    if "Summary" in wb.sheetnames:
        wb["Summary"]["B3"] = today
    wb.save(TRACKER_PATH)  # close the file in Excel before running locally
    print(f"{datetime.now():%Y-%m-%d %H:%M} - added {len(new)} new jobs "
          f"({sum(1 for n in new if tier(n['location']) == 1)} Mumbai/Remote)")


if __name__ == "__main__":
    main()
