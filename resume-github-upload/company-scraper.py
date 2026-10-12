#!/usr/bin/env python3
"""
Infopark Company Directory Scraper
------------------------------------
Scrapes the FULL company directory at infopark.in/companies (not just job
listings) — every company that has ever set up at Infopark, regardless of
whether they currently have open positions.

For each company it:
    1. Reads the paginated A-Z listing to get every company's profile URL.
    2. Visits each company's profile page and extracts the email address
       (it's embedded in the address line, e.g. "...682303 - careers@x.com").
    3. Writes Company Name + Email to Excel, batched into files of 30 rows.

This covers ~350-580 companies (varies by location), so a full run with the
default 1 second delay takes a while (roughly 15-20 minutes for all of
Kochi Phase 1). Use --pages to test on a smaller slice first.

Usage:
    python infopark_companies_scraper.py
    python infopark_companies_scraper.py --location infopark-kochi-phase-1
    python infopark_companies_scraper.py --pages 2          # quick test
    python infopark_companies_scraper.py --all-locations    # every campus

Output:
    infopark_companies_all_part1.xlsx   (30 rows: Company Name | Email)
    infopark_companies_all_part2.xlsx
    ...

Requirements:
    pip install requests beautifulsoup4 openpyxl --break-system-packages
"""

import argparse
import re
import sys
import time

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font

BASE_URL = "https://infopark.in"

LOCATIONS = {
    "all": "companies",
    "infopark-kochi-phase-1": "companies/infopark-kochi-phase-1",
    "infopark-kochi-phase-2": "companies/infopark-kochi-phase-2",
    "infopark-cherthala": "companies/infopark-cherthala",
    "infopark-thrissur": "companies/infopark-thrissur",
    "i-by-infopark-ernakulam-south-metro-station": "companies/i-by-infopark-ernakulam-south-metro-station",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")


def get_soup(url, session, retries=3, delay=1.5):
    for attempt in range(1, retries + 1):
        try:
            resp = session.get(url, headers=HEADERS, timeout=20)
            resp.raise_for_status()
            return BeautifulSoup(resp.text, "html.parser")
        except requests.RequestException as exc:
            print(f"  [warn] attempt {attempt}/{retries} failed for {url}: {exc}")
            time.sleep(delay)
    print(f"  [error] giving up on {url}")
    return None


def find_last_page(soup):
    max_page = 1
    for a in soup.select("a[href*='page=']"):
        m = re.search(r"page=(\d+)", a.get("href", ""))
        if m:
            max_page = max(max_page, int(m.group(1)))
    return max_page


def parse_directory_page(soup):
    """Each company card is an <h5>-ish heading followed by a 'Company Profile' link."""
    companies = []
    for a in soup.find_all("a", href=True):
        if "/companies-profile/" not in a["href"]:
            continue
        if a.get_text(strip=True) != "Company Profile":
            continue
        profile_url = a["href"]
        if profile_url.startswith("/"):
            profile_url = BASE_URL + profile_url

        # The company name is the nearest heading before this link.
        heading = a.find_previous(["h5", "h4", "h3"])
        name = heading.get_text(strip=True) if heading else ""
        companies.append({"company_name": name, "profile_url": profile_url})
    return companies


def parse_profile_page(soup):
    """Email is embedded in the address paragraph, e.g. '...682303 - foo@bar.com'."""
    text = soup.get_text(" ", strip=True)
    m = EMAIL_RE.search(text)
    return m.group(0) if m else ""


def scrape_directory(location_key, max_pages=None, delay=1.0):
    session = requests.Session()
    all_companies = []
    seen_urls = set()

    path = LOCATIONS.get(location_key, "companies")
    listing_url = f"{BASE_URL}/{path}"
    print(f"\n=== Scraping company directory: {listing_url} ===")
    soup = get_soup(listing_url, session)
    if soup is None:
        return all_companies

    last_page = find_last_page(soup)
    if max_pages:
        last_page = min(last_page, max_pages)
    print(f"Found {last_page} page(s) of companies.")

    for page in range(1, last_page + 1):
        page_url = listing_url if page == 1 else f"{listing_url}?page={page}"
        print(f"\n--- Page {page}/{last_page}: {page_url}")
        page_soup = soup if page == 1 else get_soup(page_url, session)
        if page_soup is None:
            continue

        rows = parse_directory_page(page_soup)
        print(f"  Found {len(rows)} company card(s) on this page.")

        for row in rows:
            if row["profile_url"] in seen_urls:
                continue
            seen_urls.add(row["profile_url"])

            profile_soup = get_soup(row["profile_url"], session)
            email = parse_profile_page(profile_soup) if profile_soup else ""
            row["email"] = email
            print(f"  -> {row['company_name']}: {email or '(no email found)'}")
            all_companies.append(row)
            time.sleep(delay)

        time.sleep(delay)

    return all_companies


def save_excel_batches(companies, base_name, batch_size=30):
    seen = set()
    pairs = []
    for c in companies:
        name = (c.get("company_name") or "").strip()
        email = (c.get("email") or "").strip()
        if not name or not email:
            continue
        key = (name.lower(), email.lower())
        if key in seen:
            continue
        seen.add(key)
        pairs.append((name, email))

    if not pairs:
        print(f"No company/email pairs to write for {base_name}.")
        return

    total_parts = (len(pairs) + batch_size - 1) // batch_size
    for part in range(total_parts):
        chunk = pairs[part * batch_size : (part + 1) * batch_size]
        wb = Workbook()
        ws = wb.active
        ws.title = "Companies"

        ws["A1"] = "Company Name"
        ws["B1"] = "Email"
        for cell in ("A1", "B1"):
            ws[cell].font = Font(name="Arial", bold=True)

        for i, (name, email) in enumerate(chunk, start=2):
            ws[f"A{i}"] = name
            ws[f"B{i}"] = email
            ws[f"A{i}"].font = Font(name="Arial")
            ws[f"B{i}"].font = Font(name="Arial")

        ws.column_dimensions["A"].width = 40
        ws.column_dimensions["B"].width = 35

        filename = f"{base_name}_part{part + 1}.xlsx"
        wb.save(filename)
        print(f"Saved {len(chunk)} row(s) to {filename}")

    print(
        f"\nTotal unique company/email pairs: {len(pairs)} "
        f"across {total_parts} Excel file(s) of up to {batch_size} rows each."
    )


def main():
    parser = argparse.ArgumentParser(description="Scrape the full Infopark company directory.")
    parser.add_argument(
        "--location",
        default="all",
        choices=list(LOCATIONS.keys()),
        help="Which directory to scrape. 'all' = every location combined (default).",
    )
    parser.add_argument(
        "--all-locations",
        action="store_true",
        help="Scrape each location separately (in addition to / instead of --location).",
    )
    parser.add_argument("--pages", type=int, default=None, help="Limit to first N pages (for testing).")
    parser.add_argument("--delay", type=float, default=1.0, help="Delay between requests (seconds).")
    parser.add_argument("--batch-size", type=int, default=30, help="Rows per Excel file.")
    args = parser.parse_args()

    if args.all_locations:
        targets = [k for k in LOCATIONS if k != "all"]
    else:
        targets = [args.location]

    for loc in targets:
        companies = scrape_directory(loc, max_pages=args.pages, delay=args.delay)
        if not companies:
            print(f"No companies found for {loc}.")
            continue
        base_name = f"infopark_companies_{loc.replace('-', '_')}"
        save_excel_batches(companies, base_name, batch_size=args.batch_size)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        sys.exit(1)