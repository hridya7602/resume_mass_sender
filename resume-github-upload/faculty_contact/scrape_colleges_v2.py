import re, time, os, requests
from bs4 import BeautifulSoup
import openpyxl
from urllib.parse import quote_plus, urlparse

# ─────────────────────────────────────────────────────────────────────────────
# CollegeDunia uses  /STATE-NAME/COURSE-colleges  (state in PATH, not ?param)
# e.g. https://collegedunia.com/kerala/ba-arts-colleges
# ─────────────────────────────────────────────────────────────────────────────

COLLEGE_LIST_URLS = [
    # ── SOUTH INDIA FIRST ────────────────────────────────────────────────────
    ("Kerala",         "https://collegedunia.com/kerala/ba-arts-colleges"),
    ("Tamil Nadu",     "https://collegedunia.com/tamil-nadu/ba-arts-colleges"),
    ("Karnataka",      "https://collegedunia.com/karnataka/ba-arts-colleges"),
    ("Andhra Pradesh", "https://collegedunia.com/andhra-pradesh/ba-arts-colleges"),
    ("Telangana",      "https://collegedunia.com/telangana/ba-arts-colleges"),
    ("Tamil Nadu",     "https://collegedunia.com/tamil-nadu/btech-colleges"),
    ("Karnataka",      "https://collegedunia.com/karnataka/btech-colleges"),
    ("Andhra Pradesh", "https://collegedunia.com/andhra-pradesh/btech-colleges"),
    ("Telangana",      "https://collegedunia.com/telangana/btech-colleges"),
    ("Karnataka",      "https://collegedunia.com/karnataka/mba-colleges"),
    ("Tamil Nadu",     "https://collegedunia.com/tamil-nadu/mba-colleges"),
    # ── NORTH / WEST ─────────────────────────────────────────────────────────
    ("Delhi",          "https://collegedunia.com/delhi/ba-arts-colleges"),
    ("Uttar Pradesh",  "https://collegedunia.com/uttar-pradesh/ba-arts-colleges"),
    ("Maharashtra",    "https://collegedunia.com/maharashtra/ba-arts-colleges"),
    ("West Bengal",    "https://collegedunia.com/west-bengal/ba-arts-colleges"),
    ("Rajasthan",      "https://collegedunia.com/rajasthan/ba-arts-colleges"),
    ("Gujarat",        "https://collegedunia.com/gujarat/ba-arts-colleges"),
    ("Madhya Pradesh", "https://collegedunia.com/madhya-pradesh/ba-arts-colleges"),
    ("Bihar",          "https://collegedunia.com/bihar/ba-arts-colleges"),
    ("Punjab",         "https://collegedunia.com/punjab/ba-arts-colleges"),
    ("Haryana",        "https://collegedunia.com/haryana/ba-arts-colleges"),
    ("Odisha",         "https://collegedunia.com/odisha/ba-arts-colleges"),
    ("Uttar Pradesh",  "https://collegedunia.com/uttar-pradesh/btech-colleges"),
    ("Maharashtra",    "https://collegedunia.com/maharashtra/btech-colleges"),
    ("Rajasthan",      "https://collegedunia.com/rajasthan/btech-colleges"),
    ("Gujarat",        "https://collegedunia.com/gujarat/btech-colleges"),
    ("Delhi",          "https://collegedunia.com/delhi/btech-colleges"),
    ("West Bengal",    "https://collegedunia.com/west-bengal/btech-colleges"),
    ("Maharashtra",    "https://collegedunia.com/maharashtra/mba-colleges"),
    ("Delhi",          "https://collegedunia.com/delhi/mba-colleges"),
    ("Uttar Pradesh",  "https://collegedunia.com/uttar-pradesh/mba-colleges"),
]

DELAY      = 1.5
SAVE_EVERY = 30
OUTPUT     = "college_direct_contacts"
SENT_FILE  = "sent_emails.txt"

SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
    "Referer": "https://www.google.com/",
})

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")

SKIP_DOMAINS = {
    "collegedunia.com", "sentry.io", "example.com", "wordpress.com",
    "wix.com", "godaddy.com", "cloudflare.com", "google.com",
    "facebook.com", "twitter.com", "linkedin.com", "instagram.com",
    "shielded-form.com", "w3.org", "schema.org", "youtube.com",
    "whatsapp.com", "telegram.org",
}
PERSONAL_DOMAINS = {
    "gmail.com", "yahoo.com", "hotmail.com", "rediffmail.com",
    "outlook.com", "ymail.com",
}

CONTACT_PATHS = [
    "", "/contact", "/contact-us", "/contactus", "/contact.html",
    "/contact-us.html", "/about", "/about-us", "/administration",
    "/principal", "/recruitment", "/careers", "/vacancy", "/hr",
    "/reach-us", "/enquiry", "/info", "/faculty",
]

# ── State keyword maps for location verification ──────────────────────────────
STATE_KEYWORDS = {
    "Kerala":         ["kerala", "kochi", "thiruvananthapuram", "kozhikode",
                       "thrissur", "ernakulam", "kannur", "kollam", "palakkad",
                       "malappuram", "calicut", "trivandrum", "alappuzha"],
    "Tamil Nadu":     ["tamil nadu", "chennai", "coimbatore", "madurai",
                       "tiruchirappalli", "trichy", "salem", "tirunelveli",
                       "vellore", "tiruppur"],
    "Karnataka":      ["karnataka", "bangalore", "bengaluru", "mysore",
                       "hubli", "mangalore", "belgaum"],
    "Andhra Pradesh": ["andhra", "visakhapatnam", "vizag", "vijayawada",
                       "guntur", "tirupati", "kurnool", "rajamahendravaram"],
    "Telangana":      ["telangana", "hyderabad", "warangal", "nizamabad",
                       "karimnagar"],
    "Delhi":          ["delhi", "new delhi"],
    "Uttar Pradesh":  ["uttar pradesh", "lucknow", "kanpur", "agra",
                       "allahabad", "prayagraj", "varanasi", "noida",
                       "ghaziabad", "meerut"],
    "Maharashtra":    ["maharashtra", "mumbai", "pune", "nagpur", "nashik",
                       "aurangabad", "thane"],
    "West Bengal":    ["west bengal", "kolkata", "calcutta", "howrah",
                       "durgapur", "asansol"],
    "Rajasthan":      ["rajasthan", "jaipur", "jodhpur", "udaipur",
                       "kota", "ajmer", "bikaner"],
    "Gujarat":        ["gujarat", "ahmedabad", "surat", "vadodara",
                       "baroda", "rajkot"],
    "Madhya Pradesh": ["madhya pradesh", "bhopal", "indore", "gwalior",
                       "jabalpur"],
    "Bihar":          ["bihar", "patna", "gaya", "bhagalpur", "muzaffarpur"],
    "Punjab":         ["punjab", "chandigarh", "ludhiana", "amritsar",
                       "jalandhar"],
    "Haryana":        ["haryana", "gurgaon", "gurugram", "faridabad",
                       "hisar", "rohtak"],
    "Odisha":         ["odisha", "bhubaneswar", "cuttack", "rourkela",
                       "berhampur"],
}


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
        if r.status_code in (403, 404, 410):
            return None
        r.raise_for_status()
        return BeautifulSoup(r.text, "html.parser")
    except Exception:
        return None


def extract_emails(soup):
    if not soup: return []
    found = []
    for a in soup.find_all("a", href=re.compile(r"^mailto:", re.I)):
        e = a["href"].replace("mailto:", "").strip().split("?")[0].lower()
        if "@" in e and e not in found:
            found.append(e)
    for e in EMAIL_RE.findall(soup.get_text(separator=" ")):
        e = e.lower().rstrip(".,;:)'\"")
        domain = e.split("@")[-1]
        if domain not in SKIP_DOMAINS and e not in found:
            found.append(e)
    return found


def best_email(emails):
    if not emails: return None
    for suffix in (".ac.in", ".edu.in", ".edu"):
        for e in emails:
            if e.endswith(suffix): return e
    for e in emails:
        if e.split("@")[-1] not in PERSONAL_DOMAINS: return e
    return emails[0]


def college_is_in_state(college_text, state):
    """Check if the college name/location text matches the expected state."""
    text_lower = college_text.lower()
    for kw in STATE_KEYWORDS.get(state, []):
        if kw in text_lower:
            return True
    return False


def get_college_website(cd_url):
    soup = fetch(cd_url)
    if not soup: return None

    # Strategy 1: keyword-matched anchor links
    kw = ("official website", "visit website", "website", "visit college",
          "college website", "official site", "visit site", "go to website")
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = a.get_text(strip=True).lower()
        if "collegedunia" in href or not href.startswith("http"): continue
        if any(k in text for k in kw):
            return href.rstrip("/")

    # Strategy 2: .ac.in / .edu.in / .edu links
    for a in soup.find_all("a", href=re.compile(
            r"https?://[^/]+\.(ac\.in|edu\.in|edu)", re.I)):
        href = a["href"]
        if "collegedunia" not in href:
            return href.rstrip("/")

    # Strategy 3: any external link that looks like a college root
    for a in soup.find_all("a", href=re.compile(
            r"https?://(?!.*collegedunia)[^/]+\.(ac\.in|edu\.in|org\.in|edu|ac\.uk)",
            re.I)):
        href = a["href"]
        if urlparse(href).path.count("/") <= 2:
            return href.rstrip("/")

    return None


def get_email_from_website(website):
    base = website.rstrip("/")
    for path in CONTACT_PATHS:
        soup = fetch(base + path)
        if soup:
            emails = extract_emails(soup)
            if emails:
                e = best_email(emails)
                if e: return e.rstrip(".,;:)'\"")
        time.sleep(0.4)
    return None


def google_search_email(college_name, state=""):
    """Try multiple queries to find an email for the college via Google."""
    queries = [
        f'"{college_name}" official email principal contact',
        f'"{college_name}" {state} email site:ac.in OR site:edu.in',
        f'"{college_name}" principal email id contact',
        f'"{college_name}" HOD English contact email',
        f'{college_name} {state} college official email',
    ]
    for q in queries:
        try:
            url = f"https://www.google.com/search?q={quote_plus(q)}&num=10"
            soup = fetch(url)
            if not soup: continue
            text = soup.get_text(separator=" ")
            emails = []
            for e in EMAIL_RE.findall(text):
                e = e.lower().rstrip(".,;:)'\"")
                dom = e.split("@")[-1]
                if dom in SKIP_DOMAINS: continue
                emails.append(e)
            if emails:
                best = best_email(emails)
                if best:
                    print(f"      [Google] {best}")
                    return best
            time.sleep(1.0)
        except Exception as ex:
            print(f"      [Google error] {ex}")
    return None


def get_email_for_college(name, website, state=""):
    if website:
        email = get_email_from_website(website)
        if email: return email
    return google_search_email(name, state)


def parse_college_list(list_url, expected_state):
    """
    Parse CollegeDunia list page.
    Only keep colleges whose name/location text matches expected_state.
    """
    soup = fetch(list_url)
    if not soup:
        print(f"  [!] Could not fetch: {list_url}")
        return []

    colleges, seen_urls = [], set()

    # Each college card usually has a link like /college/12345-name-of-college
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

        # ── KEY FIX: verify this college belongs to the expected state ──────
        # College name on CollegeDunia often ends with ",City" so check that
        if not college_is_in_state(name, expected_state):
            # Also check the parent element text which may include city/state
            parent_text = ""
            if a.parent:
                parent_text = a.parent.get_text(separator=" ")
            if not college_is_in_state(parent_text, expected_state):
                continue  # skip — belongs to a different state

        seen_urls.add(href)
        colleges.append((name, href))

    print(f"  → {len(colleges)} colleges matched state '{expected_state}' "
          f"(from page: {list_url.split('/')[-1]})")
    return colleges


class Writer:
    def __init__(self, prefix, size):
        self.prefix = prefix; self.size = size
        self.n = 1; self.count = 0
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


def main():
    sent   = load_sent()
    writer = Writer(OUTPUT, SAVE_EVERY)
    seen   = set(sent)
    saved = visited = no_web = no_email = dupes = 0

    print("South Indian states FIRST.\n")

    try:
        for state, list_url in COLLEGE_LIST_URLS:
            print(f"\n{'='*60}\nState: {state}  |  {list_url.split('/')[-1]}\n{'='*60}")
            colleges = parse_college_list(list_url, state)

            for name, cd_url in colleges:
                visited += 1
                time.sleep(DELAY)

                website = get_college_website(cd_url)
                if not website:
                    no_web += 1
                    print(f"  ~ no website : {name[:55]} → Google fallback …")

                time.sleep(DELAY)
                email = get_email_for_college(name, website, state)

                if not email:
                    no_email += 1
                    print(f"  ✗ no email   : {name[:55]}")
                    continue

                email = email.lower().rstrip(".,;:)'\"")
                if email in seen:
                    dupes += 1
                    continue

                seen.add(email)
                writer.add(name, state, email, website or "")
                saved += 1
                print(f"  ✓ {name[:50]!r:52s} | {email:40s} | {(website or '')[:35]}")

    except KeyboardInterrupt:
        print("\n[Stopped]")

    writer.flush()
    print(
        f"\n{'─'*60}\n"
        f"Visited:{visited}  Saved:{saved}  "
        f"NoWeb:{no_web}  NoEmail:{no_email}  Dupes:{dupes}"
    )


if __name__ == "__main__":
    main()