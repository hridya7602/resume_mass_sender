import re
import time
import requests
from bs4 import BeautifulSoup
import openpyxl

# ── Configuration ──────────────────────────────────────────────────────────────
BASE_URL         = "https://www.facultyplus.com/category/arts-and-science/language/english/"
MAX_PAGES        = 242
DELAY_SECONDS    = 1.0
SAVE_EVERY       = 30        # create a new Excel file every N saved contacts
OUTPUT_PREFIX    = "contacts_list"  # files: contacts_list_1.xlsx, contacts_list_2.xlsx …

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}

EMAIL_RE    = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
CF_EMAIL_RE = re.compile(r'data-cfemail="([a-f0-9]+)"')

EXCLUDE_DOMAINS = {"sentry.io", "example.com", "wordpress.com", "woocommerce.com"}

# Strips leading job-title phrases like "Assistant Professor Jobs at "
_PREFIX = re.compile(
    r"^(?:assistant\s+professor|associate\s+professor|professor|faculty|lecturer|hod|"
    r"teaching\s+staff|guest\s+teacher|teacher|staff|walk[\s\-]?in|"
    r"recruitment|job|opening|wanted)"
    r"[\s\w,&]*?\bat\b\s*",
    re.I,
)
# Strips trailing noise after the college name
_SUFFIX = re.compile(
    r"\s*[–\-|]\s*(?:faculty\s*plus.*|facultyplus.*)$"
    r"|\s*\bwanted\b.*$"
    r"|\s*\bjobs?\b.*$"
    r"|\s*\brecruitment\b.*$"
    r"|\s*\bopenings?\b.*$"
    r"|\s*\bwalk[\s\-]?in\b.*$"
    r"|\s*[\-–]\s*\bfacult\w*\b.*$"
    r"|\s*\bfacult\w*\b\s*$"
    r"|\s*\d{4}\s*(?:facult\w*)?\s*$",
    re.I,
)


# ── Helpers ────────────────────────────────────────────────────────────────────

def decode_cf_email(encoded: str) -> str:
    r = int(encoded[:2], 16)
    return "".join(chr(int(encoded[n:n + 2], 16) ^ r) for n in range(2, len(encoded), 2))


def get_page(url: str):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        if resp.status_code == 404:
            return None
        resp.raise_for_status()
        return BeautifulSoup(resp.text, "html.parser")
    except Exception as exc:
        print(f"  [WARN] Could not fetch {url}: {exc}")
        return None


def listing_page_url(page: int) -> str:
    return BASE_URL if page == 1 else f"{BASE_URL}page/{page}/"


def collect_detail_urls(soup) -> list:
    urls = []
    for tag in soup.select("h2.entry-title a, h3.entry-title a, article.post a.entry-title-link"):
        href = tag.get("href", "")
        if href.startswith("https://www.facultyplus.com/") and href not in urls:
            urls.append(href)
    if not urls:
        for tag in soup.select("article h2 a, article h3 a"):
            href = tag.get("href", "")
            if href.startswith("https://www.facultyplus.com/") and href not in urls:
                urls.append(href)
    return urls


def clean_college_name(raw: str) -> str:
    name = _PREFIX.sub("", raw).strip()
    name = _SUFFIX.sub("", name).strip()
    name = re.sub(r"[\s\-–|(]+$", "", name).strip()  # trailing punctuation
    return re.sub(r"\s{2,}", " ", name)


def extract_college_name(soup) -> str:
    h1 = soup.find("h1")
    raw = h1.get_text(strip=True) if h1 else ""
    if not raw:
        title = soup.find("title")
        if title:
            raw = re.sub(r"\s*[-|]\s*Faculty.*$", "", title.get_text(strip=True), flags=re.I)
    return clean_college_name(raw)


def extract_emails(soup) -> list:
    found = []
    for encoded in CF_EMAIL_RE.findall(str(soup)):
        email = decode_cf_email(encoded)
        if "@" in email and email not in found:
            found.append(email)
    for email in EMAIL_RE.findall(soup.get_text(separator=" ")):
        domain = email.split("@")[-1].lower()
        if domain not in EXCLUDE_DOMAINS and not domain.startswith("facultyplus"):
            if email not in found:
                found.append(email)
    return found


def scrape_detail(url: str):
    soup = get_page(url)
    if not soup:
        return None
    college = extract_college_name(soup)
    emails  = extract_emails(soup)
    if not emails:
        return None
    chosen = emails[0]
    for e in emails:
        if any(e.split("@")[-1].lower().endswith(s) for s in (".ac.in", ".edu", ".edu.in")):
            chosen = e
            break
    return college, chosen


# ── Excel batch manager ────────────────────────────────────────────────────────

class BatchWriter:
    def __init__(self, prefix: str, batch_size: int):
        self.prefix     = prefix
        self.batch_size = batch_size
        self.batch_num  = 1
        self.in_batch   = 0
        self.wb = None
        self.ws = None
        self._new_batch()

    def _new_batch(self):
        self.wb = openpyxl.Workbook()
        self.ws = self.wb.active
        self.ws.append(["College Name", "Email ID"])
        self.in_batch = 0

    def _filename(self) -> str:
        return f"{self.prefix}_{self.batch_num}.xlsx"

    def add(self, college: str, email: str):
        self.ws.append([college, email])
        self.in_batch += 1
        self.wb.save(self._filename())     # save immediately — no data loss on Ctrl+C

        if self.in_batch >= self.batch_size:
            print(f"  [Batch {self.batch_num} complete → {self._filename()}]")
            self.batch_num += 1
            self._new_batch()

    def flush(self):
        if self.in_batch > 0:
            self.wb.save(self._filename())
            print(f"  [Batch {self.batch_num} saved → {self._filename()} ({self.in_batch} contacts)]")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    writer        = BatchWriter(OUTPUT_PREFIX, SAVE_EVERY)
    total_saved   = 0
    total_visited = 0

    print(f"Scraping up to {MAX_PAGES} pages · saving every {SAVE_EVERY} contacts\n")

    try:
        for page_num in range(1, MAX_PAGES + 1):
            url = listing_page_url(page_num)
            print(f"[Page {page_num}] {url}")
            soup = get_page(url)
            if not soup:
                print(f"  → Page returned 404 — stopping.")
                break

            detail_urls = collect_detail_urls(soup)
            if not detail_urls:
                print(f"  → No listings found — stopping.")
                break

            print(f"  Found {len(detail_urls)} listings")

            for detail_url in detail_urls:
                total_visited += 1
                time.sleep(DELAY_SECONDS)
                result = scrape_detail(detail_url)
                if result:
                    college, email = result
                    writer.add(college, email)
                    total_saved += 1
                    print(f"    ✓ {college[:65]!r}  →  {email}")
                else:
                    print(f"    – No email: {detail_url}")

            time.sleep(DELAY_SECONDS)

    except KeyboardInterrupt:
        print("\n[Interrupted by user]")

    writer.flush()
    print(f"\nDone. Visited {total_visited} posts, saved {total_saved} contacts across {writer.batch_num} file(s).")


if __name__ == "__main__":
    main()
