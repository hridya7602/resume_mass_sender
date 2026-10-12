#!/usr/bin/env python3
"""
BeBee Job Listings Scraper
---------------------------
IMPORTANT: Unlike infopark.in, BeBee is a job AGGREGATOR (it pulls listings
from LinkedIn, Talent.com, ZipRecruiter, TheirStack, etc.). Job detail pages
only have an "Apply now" button — there is NO company email published
anywhere on the page. So this script extracts what's actually available:

    - Job Title
    - Company Name
    - Location
    - Employment Type / Work Mode (if shown)
    - Posted date
    - Job description snippet
    - Link to the job page on BeBee (which has an "Apply" button)

If you need an email, you'd have to click "Apply" on each listing yourself —
BeBee routes that to the original source (company career page, LinkedIn,
ZipRecruiter, etc.), and only THAT final destination might show an email.

Usage:
    python bebee_scraper.py
    python bebee_scraper.py --role software-engineer --city kochi
    python bebee_scraper.py --role software-engineer --city kochi --pages 3
    python bebee_scraper.py --url "https://bebee.com/in/jobs/role/software-engineer/kochi"

Output:
    bebee_software-engineer_kochi_part1.xlsx  (30 rows)
    bebee_software-engineer_kochi_part2.xlsx  (30 rows)
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

BASE_URL = "https://bebee.com"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}


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
    """BeBee pagination links use ?page=0,1,2... — find the highest one."""
    max_page = 0
    for a in soup.select("a[href*='page=']"):
        m = re.search(r"page=(\d+)", a.get("href", ""))
        if m:
            max_page = max(max_page, int(m.group(1)))
    return max_page  # 0-indexed, so this is the last page NUMBER, add 1 for count


def parse_listing_page(soup):
    """
    Each job card is an <h3> with a link to the job detail page, followed by
    the company name/link, location, and a short description snippet.
    """
    jobs = []
    for h3 in soup.find_all("h3"):
        link = h3.find("a", href=True)
        if not link or "/jobs/" not in link["href"]:
            continue

        title = link.get_text(strip=True)
        detail_url = link["href"]
        if detail_url.startswith("/"):
            detail_url = BASE_URL + detail_url

        # Walk forward through siblings of the card container to pick up
        # company name, location, and snippet text near this heading.
        card = h3.find_parent(["article", "div", "li"])
        company = ""
        location = ""
        snippet = ""

        if card:
            # Company is usually the next link after the title link.
            company_link = None
            for a in card.find_all("a", href=True):
                if a is link:
                    continue
                if "/companies/" in a["href"] or a.get_text(strip=True):
                    company_link = a
                    break
            if company_link:
                company = company_link.get_text(strip=True)

            card_text = card.get_text(" ", strip=True)
            snippet = card_text[:300]

        jobs.append(
            {
                "job_title": title,
                "company_name": company,
                "detail_url": detail_url,
                "snippet": snippet,
            }
        )
    return jobs


def parse_detail_page(soup):
    """Pull a cleaner company name / location / posted-until date from the detail page."""
    company = ""
    location = ""

    # Company name often appears right after the logo, as a link to /companies/...
    company_link = soup.find("a", href=re.compile(r"/companies/"))
    if company_link:
        company = company_link.get_text(strip=True)

    return company, location


def scrape_bebee(role, city, max_pages=None, delay=1.0, direct_url=None):
    session = requests.Session()
    all_jobs = []

    listing_url = direct_url or f"{BASE_URL}/in/jobs/role/{role}/{city}"
    print(f"\n=== Scraping BeBee: {listing_url} ===")
    soup = get_soup(listing_url, session)
    if soup is None:
        return all_jobs

    last_page_index = find_last_page(soup)  # 0-indexed
    total_pages = last_page_index + 1
    if max_pages:
        total_pages = min(total_pages, max_pages)
    print(f"Found {total_pages} page(s) of listings.")

    seen_urls = set()
    for page in range(total_pages):
        page_url = listing_url if page == 0 else f"{listing_url}?page={page}"
        print(f"\n--- Page {page + 1}/{total_pages}: {page_url}")
        page_soup = soup if page == 0 else get_soup(page_url, session)
        if page_soup is None:
            continue

        rows = parse_listing_page(page_soup)
        print(f"  Found {len(rows)} job(s) on this page.")

        for row in rows:
            if row["detail_url"] in seen_urls:
                continue
            seen_urls.add(row["detail_url"])
            print(f"  -> {row['job_title']} @ {row['company_name'] or '(unknown company)'}")
            all_jobs.append(row)

        time.sleep(delay)

    return all_jobs


def save_excel_batches(jobs, base_name, batch_size=30):
    """
    Company Name + Apply Link (no email available from BeBee — see script
    docstring). Split into files of `batch_size` rows each.
    """
    rows = []
    for job in jobs:
        company = (job.get("company_name") or "").strip()
        title = (job.get("job_title") or "").strip()
        url = (job.get("detail_url") or "").strip()
        if not company and not title:
            continue
        rows.append((company, title, url))

    if not rows:
        print(f"No rows to write for {base_name}.")
        return

    total_parts = (len(rows) + batch_size - 1) // batch_size
    for part in range(total_parts):
        chunk = rows[part * batch_size : (part + 1) * batch_size]
        wb = Workbook()
        ws = wb.active
        ws.title = "Jobs"

        ws["A1"] = "Company Name"
        ws["B1"] = "Job Title"
        ws["C1"] = "Apply Link (BeBee)"
        for cell in ("A1", "B1", "C1"):
            ws[cell].font = Font(name="Arial", bold=True)

        for i, (company, title, url) in enumerate(chunk, start=2):
            ws[f"A{i}"] = company
            ws[f"B{i}"] = title
            ws[f"C{i}"] = url
            for col in ("A", "B", "C"):
                ws[f"{col}{i}"].font = Font(name="Arial")

        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 45
        ws.column_dimensions["C"].width = 60

        filename = f"{base_name}_part{part + 1}.xlsx"
        wb.save(filename)
        print(f"Saved {len(chunk)} row(s) to {filename}")

    print(f"\nTotal rows: {len(rows)} across {total_parts} Excel file(s).")


def main():
    parser = argparse.ArgumentParser(description="Scrape BeBee job listings.")
    parser.add_argument("--role", default="software-engineer", help="Role slug, e.g. software-engineer")
    parser.add_argument("--city", default="kochi", help="City slug, e.g. kochi")
    parser.add_argument("--url", default=None, help="Full listing URL instead of --role/--city")
    parser.add_argument("--pages", type=int, default=None, help="Limit to first N pages")
    parser.add_argument("--delay", type=float, default=1.0, help="Delay between requests (seconds)")
    parser.add_argument("--batch-size", type=int, default=30, help="Rows per Excel file")
    args = parser.parse_args()

    jobs = scrape_bebee(
        role=args.role,
        city=args.city,
        max_pages=args.pages,
        delay=args.delay,
        direct_url=args.url,
    )

    if not jobs:
        print("No jobs found.")
        return

    base_name = f"bebee_{args.role}_{args.city}".replace("/", "_")
    save_excel_batches(jobs, base_name, batch_size=args.batch_size)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        sys.exit(1)