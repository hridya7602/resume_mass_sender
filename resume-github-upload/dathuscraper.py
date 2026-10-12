#!/usr/bin/env python3
"""
Infopark Job Listings Scraper
------------------------------
Scrapes job listings from infopark.in across all paginated pages for a given
location, visits each job's detail page, and extracts Company Name + Email
(deduplicated), then writes them to Excel files in batches of 30 rows each
(part1, part2, part3, ...).

Usage:
    python infopark_scraper.py
    python infopark_scraper.py --location infopark-kochi-phase-2
    python infopark_scraper.py --location infopark-kochi-phase-1 --pages 5
    python infopark_scraper.py --all-locations
    python infopark_scraper.py --batch-size 30
    python infopark_scraper.py --csv-also   # also dump the full CSV with all fields

Output:
    infopark_kochi_phase_1_part1.xlsx   (rows 1-30:  Company Name | Email)
    infopark_kochi_phase_1_part2.xlsx   (rows 31-60: Company Name | Email)
    ... one file per 30 rows ...

Requirements:
    pip install requests beautifulsoup4 openpyxl --break-system-packages
"""

import argparse
import csv
import re
import sys
import time

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font

BASE_URL = "https://infopark.in"

LOCATIONS = [
    "infopark-kochi-phase-1",
    "infopark-kochi-phase-2",
    "infopark-cherthala",
    "infopark-thrissur",
    "i-by-infopark-ernakulam-south-metro-station",
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")


def get_soup(url, session, retries=3, delay=1.5):
    """Fetch a URL and return a BeautifulSoup object, with basic retry logic."""
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
    """Look at the pagination links and return the highest page number found."""
    max_page = 1
    for a in soup.select("a[href*='page=']"):
        m = re.search(r"page=(\d+)", a.get("href", ""))
        if m:
            max_page = max(max_page, int(m.group(1)))
    return max_page


def parse_listing_page(soup):
    """
    From a listing page, extract each row's Details link plus the
    date/title/company/last-date columns shown in the table.
    Returns a list of dicts.
    """
    rows = []
    table = soup.find("table")
    if not table:
        return rows

    for tr in table.find_all("tr"):
        cells = tr.find_all("td")
        if len(cells) < 4:
            continue
        link_tag = tr.find("a", href=True)
        if not link_tag or "/company-jobs/details/" not in link_tag["href"]:
            continue

        detail_url = link_tag["href"]
        if detail_url.startswith("/"):
            detail_url = BASE_URL + detail_url

        rows.append(
            {
                "date_posted": cells[0].get_text(strip=True),
                "job_title": cells[1].get_text(strip=True),
                "company_name": cells[2].get_text(strip=True),
                "last_date": cells[3].get_text(strip=True),
                "detail_url": detail_url,
            }
        )
    return rows


def parse_detail_page(soup):
    """
    From a job detail page, pull the company contact block near the top
    (phone + email shown right under the company name), falling back to
    scanning the whole page text for an email if needed.
    """
    phone = ""
    email = ""

    # The contact block usually sits right after the company logo/name,
    # before the "Company Profile" link. Search all text nodes for the
    # first email-looking string and a phone-looking string near the top.
    page_text = soup.get_text("\n", strip=True)
    lines = page_text.split("\n")

    email_match = EMAIL_RE.search(page_text)
    if email_match:
        email = email_match.group(0)

    phone_re = re.compile(r"(\+?\d[\d\-\s]{7,14}\d)")
    for line in lines[:40]:  # phone is near the top, close to company name
        m = phone_re.search(line)
        if m and "@" not in line:
            phone = m.group(1).strip()
            break

    # Prefer the "apply" email explicitly mentioned near "share your
    # updated resume" / "To Apply" / "send your resume", since some
    # listings give a job-specific address different from the general one.
    apply_context = re.search(
        r"(?:share your updated resume with us at|to apply|send your resume|send.{0,15}resume to)\s*[:\-]?\s*\**<?([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})",
        page_text,
        re.IGNORECASE,
    )
    if apply_context:
        email = apply_context.group(1)

    return email, phone


def scrape_location(location, max_pages=None, delay=1.0):
    session = requests.Session()
    all_jobs = []

    listing_url = f"{BASE_URL}/companies-job/{location}"
    print(f"\n=== Scraping location: {location} ===")
    soup = get_soup(listing_url, session)
    if soup is None:
        return all_jobs

    last_page = find_last_page(soup)
    if max_pages:
        last_page = min(last_page, max_pages)
    print(f"Found {last_page} page(s) of listings.")

    for page in range(1, last_page + 1):
        page_url = listing_url if page == 1 else f"{listing_url}?page={page}"
        print(f"\n--- Page {page}/{last_page}: {page_url}")
        page_soup = soup if page == 1 else get_soup(page_url, session)
        if page_soup is None:
            continue

        rows = parse_listing_page(page_soup)
        print(f"  Found {len(rows)} job(s) on this page.")

        for row in rows:
            print(f"  -> {row['job_title']} @ {row['company_name']}")
            detail_soup = get_soup(row["detail_url"], session)
            if detail_soup is None:
                row["email"] = ""
                row["phone"] = ""
            else:
                email, phone = parse_detail_page(detail_soup)
                row["email"] = email
                row["phone"] = phone
            all_jobs.append(row)
            time.sleep(delay)  # be polite to the server

        time.sleep(delay)

    return all_jobs


def save_csv(jobs, filename):
    fieldnames = [
        "date_posted",
        "job_title",
        "company_name",
        "email",
        "phone",
        "last_date",
        "detail_url",
    ]
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for job in jobs:
            writer.writerow({k: job.get(k, "") for k in fieldnames})
    print(f"\nSaved {len(jobs)} job(s) to {filename}")


def save_excel_batches(jobs, base_name, batch_size=30):
    """
    Write only Company Name + Email to Excel, split into multiple files of
    `batch_size` rows each (e.g. rows 1-30 -> _part1.xlsx, 31-60 -> _part2.xlsx).
    Duplicate (company, email) pairs across jobs are dropped, so each company
    appears once per file set.
    """
    seen = set()
    pairs = []
    for job in jobs:
        company = (job.get("company_name") or "").strip()
        email = (job.get("email") or "").strip()
        if not company or not email:
            continue
        key = (company.lower(), email.lower())
        if key in seen:
            continue
        seen.add(key)
        pairs.append((company, email))

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

        for i, (company, email) in enumerate(chunk, start=2):
            ws[f"A{i}"] = company
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
    parser = argparse.ArgumentParser(description="Scrape Infopark job listings.")
    parser.add_argument(
        "--location",
        default="infopark-kochi-phase-1",
        choices=LOCATIONS,
        help="Which Infopark location's job listing page to scrape.",
    )
    parser.add_argument(
        "--all-locations",
        action="store_true",
        help="Scrape every location listed in LOCATIONS, each to its own CSV.",
    )
    parser.add_argument(
        "--pages",
        type=int,
        default=None,
        help="Limit to the first N pages (default: scrape all available pages).",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.0,
        help="Delay in seconds between requests (default: 1.0). Increase to be gentler on the server.",
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        default=30,
        help="Rows per Excel file for the Company Name / Email output (default: 30).",
    )
    parser.add_argument(
        "--csv-also",
        action="store_true",
        help="Also write the full CSV (all fields) in addition to the batched Excel files.",
    )
    args = parser.parse_args()

    locations = LOCATIONS if args.all_locations else [args.location]

    for loc in locations:
        jobs = scrape_location(loc, max_pages=args.pages, delay=args.delay)
        if not jobs:
            print(f"No jobs found for {loc}.")
            continue

        base_name = loc.replace("-", "_")
        save_excel_batches(jobs, base_name, batch_size=args.batch_size)

        if args.csv_also:
            save_csv(jobs, f"{base_name}_jobs.csv")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        sys.exit(1)