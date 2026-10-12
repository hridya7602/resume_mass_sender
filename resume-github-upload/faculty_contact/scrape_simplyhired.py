import re, time, os, requests
from bs4 import BeautifulSoup
import openpyxl
from urllib.parse import quote_plus, urlparse

# ── South Indian states FIRST, then rest ─────────────────────────────────────
COLLEGE_LIST_URLS = [
    # ── SOUTH INDIA FIRST ────────────────────────────────────────────────────
    ("Kerala - BA",         "https://collegedunia.com/ba-colleges?state=kerala"),
    ("Tamil Nadu - BA",     "https://collegedunia.com/ba-colleges?state=tamil-nadu"),
    ("Karnataka - BA",      "https://collegedunia.com/ba-colleges?state=karnataka"),
    ("Andhra Pradesh - BA", "https://collegedunia.com/ba-colleges?state=andhra-pradesh"),
    ("Telangana - BA",      "https://collegedunia.com/ba-colleges?state=telangana"),
    ("Tamil Nadu - Engg",   "https://collegedunia.com/btech-colleges?state=tamil-nadu"),
    ("Karnataka - Engg",    "https://collegedunia.com/btech-colleges?state=karnataka"),
    ("AP - Engg",           "https://collegedunia.com/btech-colleges?state=andhra-pradesh"),
    ("Telangana - Engg",    "https://collegedunia.com/btech-colleges?state=telangana"),
    ("Karnataka - MBA",     "https://collegedunia.com/mba-colleges?state=karnataka"),
    ("Tamil Nadu - MBA",    "https://collegedunia.com/mba-colleges?state=tamil-nadu"),
    # ── NORTH / WEST INDIA ───────────────────────────────────────────────────
    ("Delhi - BA",          "https://collegedunia.com/ba-colleges?state=delhi"),
    ("Uttar Pradesh - BA",  "https://collegedunia.com/ba-colleges?state=uttar-pradesh"),
    ("Maharashtra - BA",    "https://collegedunia.com/ba-colleges?state=maharashtra"),
    ("West Bengal - BA",    "https://collegedunia.com/ba-colleges?state=west-bengal"),
    ("Rajasthan - BA",      "https://collegedunia.com/ba-colleges?state=rajasthan"),
    ("Gujarat - BA",        "https://collegedunia.com/ba-colleges?state=gujarat"),
    ("Madhya Pradesh - BA", "https://collegedunia.com/ba-colleges?state=madhya-pradesh"),
    ("Bihar - BA",          "https://collegedunia.com/ba-colleges?state=bihar"),
    ("Punjab - BA",         "https://collegedunia.com/ba-colleges?state=punjab"),
    ("Haryana - BA",        "https://collegedunia.com/ba-colleges?state=haryana"),
    ("Odisha - BA",         "https://collegedunia.com/ba-colleges?state=odisha"),
    ("UP - Engg",           "https://collegedunia.com/btech-colleges?state=uttar-pradesh"),
    ("Maharashtra - Engg",  "https://collegedunia.com/btech-colleges?state=maharashtra"),
    ("Rajasthan - Engg",    "https://collegedunia.com/btech-colleges?state=rajasthan"),
    ("Gujarat - Engg",      "https://collegedunia.com/btech-colleges?state=gujarat"),
    ("Delhi - Engg",        "https://collegedunia.com/btech-colleges?state=delhi"),
    ("West Bengal - Engg",  "https://collegedunia.com/btech-colleges?state=west-bengal"),
    ("Maharashtra - MBA",   "https://collegedunia.com/mba-colleges?state=maharashtra"),
    ("Delhi - MBA",         "https://collegedunia.com/mba-colleges?state=delhi"),
    ("UP - MBA",            "https://collegedunia.com/mba-colleges?state=uttar-pradesh"),
]

DELAY      = 1.5
SAVE_EVERY = 30
OUTPUT     = "college_direct_contacts"
SENT_FILE  = "sent_emails.txt"

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
    "Referer": "https://www.google.com/",
})

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
SKIP_DOMAINS = {
    "collegedunia.com", "sentry.io", "example.com", "wordpress.com",
    "wix.com", "godaddy.com", "cloudflare.com", "google.com",
    "facebook.com", "twitter.com", "linkedin.com", "instagram.com",
    "shielded-form.com", "spamdomain.com", "w3.org", "schema.org",
}
PERSONAL_DOMAINS = {"gmail.com", "yahoo.com", "hotmail.com", "rediffmail.com", "outlook.com", "ymail.com"}

CONTACT_PATHS = [
    "", "/contact", "/contact-us", "/contactus", "/contact.html",
    "/contact-us.html", "/about", "/about-us", "/administration",
    "/principal", "/recruitment", "/careers", "/vacancy", "/hr",
    "/reach-us", "/reach_us", "/enquiry", "/info",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def load_sent():
    sent = set()
    if os.path.exists(SENT_FILE):
        with open(SENT_FILE) as f:
            for line in f:
                e = line.strip().lower()
                if e: sent.add(e)
        print(f"Loaded {len(sent)} already-sent emails from {SENT_FILE}")
    return sent

def fetch(url, timeout=12):
    try:
        r = SESSION.get(url, timeout=timeout, allow_redirects=True)
        if r.status_code in (403, 404, 410): return None
        r.raise_for_status()
        return BeautifulSoup(r.text, "html.parser")
    except Exception:
        return None

def extract_emails(soup):
    if not soup: return []
    found = []
    # 1. mailto links first (most reliable)
    for a in soup.find_all("a", href=re.compile(r"^mailto:", re.I)):
        e = a["href"].replace("mailto:", "").strip().split("?")[0].lower()
        if "@" in e and e not in found:
            found.append(e)
    # 2. plain text scan
    for e in EMAIL_RE.findall(soup.get_text(separator=" ")):
        e = e.lower()
        domain = e.split("@")[-1]
        if domain not in SKIP_DOMAINS and e not in found:
            found.append(e)
    return found

def best_email(emails):
    if not emails: return None
    # Prefer institutional domains
    for suffix in (".ac.in", ".edu.in", ".edu"):
        for e in emails:
            if e.endswith(suffix): return e
    # Prefer any non-personal domain
    for e in emails:
        if e.split("@")[-1] not in PERSONAL_DOMAINS: return e
    return emails[0]

def clean_email(e):
    """Remove trailing punctuation that gets stuck to emails."""
    return e.strip().rstrip(".,;:)'\"")


# ── Website discovery: much more aggressive ───────────────────────────────────

def get_college_website(cd_url):
    soup = fetch(cd_url)
    if not soup: return None

    # Strategy 1: any link with common official-website keywords
    kw = ("official website", "visit website", "website", "visit college",
          "college website", "official site", "visit site", "go to website")
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = a.get_text(strip=True).lower()
        if "collegedunia" in href: continue
        if not href.startswith("http"): continue
        if any(k in text for k in kw):
            return href.rstrip("/")

    # Strategy 2: any .ac.in / .edu.in / .edu link on the page
    for a in soup.find_all("a", href=re.compile(r"https?://[^/]+\.(ac\.in|edu\.in|edu)", re.I)):
        href = a["href"]
        if "collegedunia" not in href:
            return href.rstrip("/")

    # Strategy 3: look inside structured info boxes / meta tags
    for meta in soup.find_all("meta", attrs={"property": re.compile(r"url|website", re.I)}):
        content = meta.get("content", "")
        if content.startswith("http") and "collegedunia" not in content:
            return content.rstrip("/")

    # Strategy 4: any external link that looks like a college homepage
    for a in soup.find_all("a", href=re.compile(r"https?://(?!.*collegedunia)[^/]+\.(ac\.in|edu\.in|org\.in|edu|ac\.uk)", re.I)):
        href = a["href"]
        parsed = urlparse(href)
        # skip deep pages, we want root or /contact style
        if parsed.path.count("/") <= 2:
            return href.rstrip("/")

    return None


# ── Email hunting: website first, then Google ─────────────────────────────────

def get_email_from_website(website):
    base = website.rstrip("/")
    for path in CONTACT_PATHS:
        soup = fetch(base + path)
        if soup:
            emails = extract_emails(soup)
            if emails:
                e = best_email(emails)
                if e: return clean_email(e)
        time.sleep(0.4)
    return None


def google_search_email(college_name, state=""):
    """
    Try multiple Google queries to find a direct email for the college.
    Returns best email found or None.
    """
    queries = [
        f'"{college_name}" official email principal contact',
        f'"{college_name}" {state} email "@ac.in" OR "@edu.in"',
        f'site:{college_name.lower().replace(" ","")}* email',   # approximate
        f'"{college_name}" principal email id',
        f'"{college_name}" HOD English email',
    ]
    for q in queries:
        try:
            url = f"https://www.google.com/search?q={quote_plus(q)}&num=10"
            soup = fetch(url)
            if not soup: continue
            text = soup.get_text(separator=" ")
            emails = []
            for e in EMAIL_RE.findall(text):
                e = clean_email(e.lower())
                dom = e.split("@")[-1]
                if dom in SKIP_DOMAINS: continue
                emails.append(e)
            if emails:
                best = best_email(emails)
                if best:
                    print(f"      [Google hit] {best}  (query: {q[:60]})")
                    return best
            time.sleep(0.8)
        except Exception as ex:
            print(f"      [Google error] {ex}")
    return None


def get_email_for_college(college_name, website, state=""):
    """
    Full pipeline:
      1. Scrape website contact pages
      2. If no website OR no email → Google search
    """
    if website:
        email = get_email_from_website(website)
        if email:
            return email
    # Fallback: Google regardless of whether we have a website
    return google_search_email(college_name, state)


# ── CollegeDunia list parser ──────────────────────────────────────────────────

def parse_college_list(list_url):
    soup = fetch(list_url)
    if not soup: return []
    colleges, seen_urls = [], set()
    for a in soup.find_all("a", href=re.compile(r"^/college/\d+")):
        href = "https://collegedunia.com" + a["href"]
        if href in seen_urls: continue
        name = a.get_text(strip=True)
        if not name or len(name) < 6: continue
        if re.search(
            r"[₹%#]|\bfees?\b|\bpackage\b|\branking\b|\breviews?\b"
            r"|\bbased\b|\buser\b|\d+th\b|in india|\d+\.\d+",
            name, re.I
        ): continue
        if not re.search(r"[a-zA-Z]{4,}", name): continue
        seen_urls.add(href)
        colleges.append((name, href))
    return colleges


# ── Excel writer ──────────────────────────────────────────────────────────────

class Writer:
    def __init__(self, prefix, size):
        self.prefix = prefix; self.size = size; self.n = 1; self.count = 0
        self.wb = self.ws = None; self._new()

    def _new(self):
        self.wb = openpyxl.Workbook(); self.ws = self.wb.active
        self.ws.append(["College Name", "State", "Email ID", "Website"])
        self.count = 0

    def _fname(self): return f"{self.prefix}_{self.n}.xlsx"

    def add(self, name, state, email, website):
        self.ws.append([name, state, email, website or ""])
        self.count += 1
        self.wb.save(self._fname())
        if self.count >= self.size:
            print(f"  [Batch {self.n} done → {self._fname()}]")
            self.n += 1; self._new()

    def flush(self):
        if self.count:
            self.wb.save(self._fname())
            print(f"  [Saved → {self._fname()} ({self.count} contacts)]")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    sent  = load_sent()
    writer = Writer(OUTPUT, SAVE_EVERY)
    seen  = set(sent)
    saved = visited = no_web = no_email = dupes = 0

    print(f"South Indian states listed FIRST.")
    print(f"Dedup: {len(sent)} already-sent emails\n")

    try:
        for state, list_url in COLLEGE_LIST_URLS:
            print(f"\n{'='*55}\nState: {state}\n{'='*55}")
            colleges = parse_college_list(list_url)
            print(f"Found {len(colleges)} colleges")

            for name, cd_url in colleges:
                visited += 1
                time.sleep(DELAY)

                # Step 1: find official website
                website = get_college_website(cd_url)
                if not website:
                    no_web += 1
                    print(f"  ~ no website  : {name[:55]}  → trying Google …")

                # Step 2: get email (website scrape + Google fallback)
                time.sleep(DELAY)
                email = get_email_for_college(name, website, state)

                if not email:
                    no_email += 1
                    print(f"  ✗ no email    : {name[:55]}")
                    continue

                email = clean_email(email)

                if email in seen:
                    dupes += 1
                    print(f"  ~ dupe        : {email}")
                    continue

                seen.add(email)
                writer.add(name, state, email, website or "")
                saved += 1
                ws = (website or "")[:40]
                print(f"  ✓ {name[:48]!r:50s} | {email} | {ws}")

    except KeyboardInterrupt:
        print("\n[Stopped by user]")

    writer.flush()
    print(
        f"\n{'─'*55}\n"
        f"Visited:{visited}  Saved:{saved}  "
        f"NoWeb:{no_web}  NoEmail:{no_email}  Dupes:{dupes}"
    )


if __name__ == "__main__":
    main()