import re, time, os, requests
from bs4 import BeautifulSoup
import openpyxl
from urllib.parse import quote_plus, urlparse

# ─────────────────────────────────────────────────────────────────────────────
# STRATEGY:
#  - For south Indian states: use hardcoded seed list of known college websites
#    (collegedunia returns 0 Kerala/south colleges via state param)
#  - For all states: also scrape collegedunia with pagination (pages 1-5)
#    and filter by state keywords
#  - For every college: scrape website contact pages → Google fallback
#  - Salary filter: skip Mahavidyalayas and unaided plain degree colleges
# ─────────────────────────────────────────────────────────────────────────────

# ── Hardcoded seed: South Indian colleges with known websites ─────────────────
# Format: (college_name, state, website_url)
SEED_COLLEGES = [
    # ── KERALA ───────────────────────────────────────────────────────────────
    ("St. Albert's College, Ernakulam",          "Kerala", "https://www.stalberts.edu.in"),
    ("St. Thomas College, Thrissur",             "Kerala", "https://www.stthomascollege.ac.in"),
    ("Sacred Heart College, Thevara",            "Kerala", "https://www.shcollege.ac.in"),
    ("Maharaja's College, Ernakulam",            "Kerala", "https://maharajascollege.edu.in"),
    ("Govt Victoria College, Palakkad",          "Kerala", "https://victoriacollege.edu.in"),
    ("St. Teresa's College, Ernakulam",          "Kerala", "https://www.teresas.ac.in"),
    ("Mar Ivanios College, Thiruvananthapuram",  "Kerala", "https://marivanios.edu.in"),
    ("St. Xavier's College, Thumba",             "Kerala", "https://www.stxaviersthumba.edu.in"),
    ("Farook College, Kozhikode",                "Kerala", "https://www.farookcollege.ac.in"),
    ("Government College, Kasaragod",            "Kerala", "https://gckvkd.ac.in"),
    ("Providence Women's College, Calicut",      "Kerala", "https://www.providencecollegecalicut.ac.in"),
    ("Zamorin's Guruvayurappan College, Calicut","Kerala", "https://zgcollege.ac.in"),
    ("MES College, Marampally",                  "Kerala", "https://www.mescollege.ac.in"),
    ("SNGS College, Pattambi",                   "Kerala", "https://www.sngscollege.ac.in"),
    ("Bishop Moore College, Mavelikara",         "Kerala", "https://www.bishopmoorecollege.org"),
    ("SN College, Kollam",                       "Kerala", "https://www.sncollegekollam.ac.in"),
    ("St. Berchmans College, Changanacherry",    "Kerala", "https://www.sbcollege.ac.in"),
    ("CMS College, Kottayam",                    "Kerala", "https://www.cmscollege.ac.in"),
    ("Baselius College, Kottayam",               "Kerala", "https://baseliuscollege.ac.in"),
    ("Assumption College, Changanacherry",       "Kerala", "https://assumptioncollege.in"),
    ("BCM College, Kottayam",                    "Kerala", "https://www.bcmcollege.ac.in"),
    ("Deva Matha College, Kuravilangad",         "Kerala", "https://www.devamathacollege.ac.in"),
    ("St. Dominic's College, Kanjirappally",     "Kerala", "https://stdominicscollege.in"),
    ("Rajagiri College of Social Sciences",      "Kerala", "https://www.rajagiri.edu"),
    ("Christ College, Irinjalakuda",             "Kerala", "https://www.christcollegeirinjalakuda.ac.in"),
    ("St. Paul's College, Kalamassery",          "Kerala", "https://www.stpaulscollege.ac.in"),
    ("Govt College, Chittur",                    "Kerala", "https://govtcollegechittur.ac.in"),
    ("Nirmalagiri College, Kuthuparamba",        "Kerala", "https://nirmalgiricollege.ac.in"),
    ("Malabar Christian College, Calicut",       "Kerala", "https://www.mccollege.ac.in"),
    ("Union Christian College, Aluva",           "Kerala", "https://www.uccollege.edu.in"),
    # ── TAMIL NADU ───────────────────────────────────────────────────────────
    ("Loyola College, Chennai",                  "Tamil Nadu", "https://www.loyolacollege.edu"),
    ("Presidency College, Chennai",              "Tamil Nadu", "https://www.presidencychennai.ac.in"),
    ("Madras Christian College, Chennai",        "Tamil Nadu", "https://mcc.edu.in"),
    ("Women's Christian College, Chennai",       "Tamil Nadu", "https://www.wcc.edu.in"),
    ("Stella Maris College, Chennai",            "Tamil Nadu", "https://www.stellamariscollege.edu.in"),
    ("Ethiraj College for Women, Chennai",       "Tamil Nadu", "https://www.ethirajcollege.in"),
    ("PSG College of Arts & Science, Coimbatore","Tamil Nadu", "https://www.psgcas.ac.in"),
    ("Govt Arts College, Coimbatore",            "Tamil Nadu", "https://www.govtartscollegecoimbatore.ac.in"),
    ("Scott Christian College, Nagercoil",       "Tamil Nadu", "https://www.scottchristiancollege.com"),
    ("Sourashtra College, Madurai",              "Tamil Nadu", "https://www.sourashtra.net"),
    ("American College, Madurai",                "Tamil Nadu", "https://www.americancollege.edu.in"),
    ("Lady Doak College, Madurai",               "Tamil Nadu", "https://www.ladydoak.edu.in"),
    ("Holy Cross College, Trichy",               "Tamil Nadu", "https://www.hcctrichy.ac.in"),
    ("National College, Trichy",                 "Tamil Nadu", "https://nationalcollegetrichy.ac.in"),
    ("Bishop Heber College, Trichy",             "Tamil Nadu", "https://www.bishopheber.ac.in"),
    ("Jamal Mohamed College, Trichy",            "Tamil Nadu", "https://www.jmc.edu"),
    ("Periyar E.V.R. College, Trichy",           "Tamil Nadu", "https://www.periyarevrcollegetrichy.ac.in"),
    ("Govt Arts College, Salem",                 "Tamil Nadu", "https://www.gacsalem.ac.in"),
    ("SRM Arts and Science College, Chennai",    "Tamil Nadu", "https://www.srmartsscience.ac.in"),
    ("Voorhees College, Vellore",                "Tamil Nadu", "https://www.voorheescollege.ac.in"),
    # ── KARNATAKA ────────────────────────────────────────────────────────────
    ("Christ University, Bangalore",             "Karnataka", "https://christuniversity.in"),
    ("St. Joseph's College of Arts & Science",   "Karnataka", "https://www.sjc.ac.in"),
    ("Mount Carmel College, Bangalore",          "Karnataka", "https://www.mccblr.edu.in"),
    ("Jyoti Nivas College, Bangalore",           "Karnataka", "https://jyotinivascollege.edu.in"),
    ("Bangalore University, Bangalore",          "Karnataka", "https://www.bangaloreuniversity.ac.in"),
    ("Mysore University, Mysore",                "Karnataka", "https://uni-mysore.ac.in"),
    ("Mangalore University, Mangalore",          "Karnataka", "https://mangaloreuniversity.ac.in"),
    ("St. Agnes College, Mangalore",             "Karnataka", "https://stagnes.edu.in"),
    ("St. Aloysius College, Mangalore",          "Karnataka", "https://www.staloysius.edu.in"),
    ("Vijaya College, Bangalore",                "Karnataka", "https://www.vijayacollege.ac.in"),
    ("Maharaja's College, Mysore",               "Karnataka", "https://www.maharajascollege.edu.in"),
    ("JSS College of Arts, Commerce & Science",  "Karnataka", "https://www.jsscacs.edu.in"),
    # ── ANDHRA PRADESH ───────────────────────────────────────────────────────
    ("Andhra University, Visakhapatnam",         "Andhra Pradesh", "https://www.andhrauniversity.edu.in"),
    ("Acharya Nagarjuna University, Guntur",     "Andhra Pradesh", "https://www.nagarjunauniversity.ac.in"),
    ("Sri Venkateswara University, Tirupati",    "Andhra Pradesh", "https://www.svuniversity.edu.in"),
    ("Govt Degree College, Vizag",               "Andhra Pradesh", "https://www.gdcvizag.ac.in"),
    ("Krishna University, Machilipatnam",        "Andhra Pradesh", "https://www.krishnauniversity.ac.in"),
    ("Rayalaseema University, Kurnool",          "Andhra Pradesh", "https://www.ruk.ac.in"),
    ("Vikrama Simhapuri University, Nellore",    "Andhra Pradesh", "https://www.simhapuriuniv.ac.in"),
    # ── TELANGANA ────────────────────────────────────────────────────────────
    ("Osmania University, Hyderabad",            "Telangana", "https://www.osmania.ac.in"),
    ("Hyderabad Central University",             "Telangana", "https://www.uohyd.ac.in"),
    ("Kakatiya University, Warangal",            "Telangana", "https://www.kakatiya.ac.in"),
    ("Nizam College, Hyderabad",                 "Telangana", "https://www.nizamcollege.ac.in"),
    ("Govt Degree College, Secunderabad",        "Telangana", "https://gdcsecunderabad.ac.in"),
    ("St. Francis College for Women, Hyderabad", "Telangana", "https://www.stfranciscollege.ac.in"),
    ("Bhavan's Vivekananda College, Hyderabad",  "Telangana", "https://www.bvvcollege.org"),
]

# ── CollegeDunia: paginated scraping for all states ──────────────────────────
PAGES = 5  # scrape up to 5 pages per state

COLLEGE_LIST_URLS = [
    # South first
    ("Kerala",         "https://collegedunia.com/ba-colleges?state=kerala"),
    ("Tamil Nadu",     "https://collegedunia.com/ba-colleges?state=tamil-nadu"),
    ("Karnataka",      "https://collegedunia.com/ba-colleges?state=karnataka"),
    ("Andhra Pradesh", "https://collegedunia.com/ba-colleges?state=andhra-pradesh"),
    ("Telangana",      "https://collegedunia.com/ba-colleges?state=telangana"),
    ("Tamil Nadu",     "https://collegedunia.com/btech-colleges?state=tamil-nadu"),
    ("Karnataka",      "https://collegedunia.com/btech-colleges?state=karnataka"),
    ("Andhra Pradesh", "https://collegedunia.com/btech-colleges?state=andhra-pradesh"),
    ("Telangana",      "https://collegedunia.com/btech-colleges?state=telangana"),
    ("Karnataka",      "https://collegedunia.com/mba-colleges?state=karnataka"),
    ("Tamil Nadu",     "https://collegedunia.com/mba-colleges?state=tamil-nadu"),
    # North
    ("Delhi",          "https://collegedunia.com/ba-colleges?state=delhi"),
    ("Uttar Pradesh",  "https://collegedunia.com/ba-colleges?state=uttar-pradesh"),
    ("Maharashtra",    "https://collegedunia.com/ba-colleges?state=maharashtra"),
    ("West Bengal",    "https://collegedunia.com/ba-colleges?state=west-bengal"),
    ("Rajasthan",      "https://collegedunia.com/ba-colleges?state=rajasthan"),
    ("Gujarat",        "https://collegedunia.com/ba-colleges?state=gujarat"),
    ("Madhya Pradesh", "https://collegedunia.com/ba-colleges?state=madhya-pradesh"),
    ("Bihar",          "https://collegedunia.com/ba-colleges?state=bihar"),
    ("Punjab",         "https://collegedunia.com/ba-colleges?state=punjab"),
    ("Haryana",        "https://collegedunia.com/ba-colleges?state=haryana"),
    ("Odisha",         "https://collegedunia.com/ba-colleges?state=odisha"),
    ("Uttar Pradesh",  "https://collegedunia.com/btech-colleges?state=uttar-pradesh"),
    ("Maharashtra",    "https://collegedunia.com/btech-colleges?state=maharashtra"),
    ("Rajasthan",      "https://collegedunia.com/btech-colleges?state=rajasthan"),
    ("Gujarat",        "https://collegedunia.com/btech-colleges?state=gujarat"),
    ("Delhi",          "https://collegedunia.com/btech-colleges?state=delhi"),
    ("West Bengal",    "https://collegedunia.com/btech-colleges?state=west-bengal"),
    ("Maharashtra",    "https://collegedunia.com/mba-colleges?state=maharashtra"),
    ("Delhi",          "https://collegedunia.com/mba-colleges?state=delhi"),
    ("Uttar Pradesh",  "https://collegedunia.com/mba-colleges?state=uttar-pradesh"),
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
    "collegedunia.com","sentry.io","example.com","wordpress.com","wix.com",
    "godaddy.com","cloudflare.com","google.com","facebook.com","twitter.com",
    "linkedin.com","instagram.com","shielded-form.com","w3.org","schema.org",
    "youtube.com","whatsapp.com","telegram.org",
}
PERSONAL_DOMAINS = {
    "gmail.com","yahoo.com","hotmail.com","rediffmail.com","outlook.com","ymail.com"
}
CONTACT_PATHS = [
    "","/ contact","/contact","/contact-us","/contactus","/contact.html",
    "/contact-us.html","/about","/about-us","/administration","/principal",
    "/recruitment","/careers","/vacancy","/hr","/reach-us","/enquiry","/info","/faculty",
]
STATE_KEYWORDS = {
    "Kerala":         ["kerala","kochi","thiruvananthapuram","kozhikode","thrissur",
                       "ernakulam","kannur","kollam","palakkad","malappuram",
                       "calicut","trivandrum","alappuzha","kottayam","idukki",
                       "mavelikara","changanacherry","irinjalakuda","aluva",
                       "pattambi","kuravilangad","kanjirappally","kuthuparamba",
                       "kalamassery","thumba","kasaragod"],
    "Tamil Nadu":     ["tamil nadu","tamil","chennai","coimbatore","madurai",
                       "tiruchirappalli","trichy","salem","tirunelveli","vellore",
                       "tiruppur","tanjore","thanjavur","chidambaram","nagercoil",
                       "madurai","kumbakonam"],
    "Karnataka":      ["karnataka","bangalore","bengaluru","mysore","mysuru",
                       "hubli","mangalore","belgaum","dharwad","tumkur","shimoga",
                       "udupi","manipal"],
    "Andhra Pradesh": ["andhra","visakhapatnam","vizag","vijayawada","guntur",
                       "tirupati","kurnool","rajamahendravaram","rajahmundry",
                       "nellore","kakinada","ongole","machilipatnam"],
    "Telangana":      ["telangana","hyderabad","warangal","nizamabad",
                       "karimnagar","khammam","secunderabad","medak"],
    "Delhi":          ["delhi","new delhi"],
    "Uttar Pradesh":  ["uttar pradesh","lucknow","kanpur","agra","allahabad",
                       "prayagraj","varanasi","noida","ghaziabad","meerut",
                       "bareilly","aligarh","gorakhpur"],
    "Maharashtra":    ["maharashtra","mumbai","pune","nagpur","nashik",
                       "aurangabad","thane","solapur","kolhapur","amravati"],
    "West Bengal":    ["west bengal","kolkata","calcutta","howrah","durgapur",
                       "asansol","siliguri","kharagpur","burdwan"],
    "Rajasthan":      ["rajasthan","jaipur","jodhpur","udaipur","kota",
                       "ajmer","bikaner","alwar","bhilwara"],
    "Gujarat":        ["gujarat","ahmedabad","surat","vadodara","baroda",
                       "rajkot","gandhinagar","bhavnagar","jamnagar"],
    "Madhya Pradesh": ["madhya pradesh","bhopal","indore","gwalior","jabalpur",
                       "ujjain","sagar","rewa"],
    "Bihar":          ["bihar","patna","gaya","bhagalpur","muzaffarpur","darbhanga"],
    "Punjab":         ["punjab","chandigarh","ludhiana","amritsar","jalandhar",
                       "patiala","bathinda"],
    "Haryana":        ["haryana","gurgaon","gurugram","faridabad","hisar",
                       "rohtak","ambala","panipat","karnal"],
    "Odisha":         ["odisha","bhubaneswar","cuttack","rourkela","berhampur",
                       "sambalpur","puri"],
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
        if r.status_code in (403, 404, 410): return None
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

def college_is_in_state(text, state):
    t = text.lower()
    return any(kw in t for kw in STATE_KEYWORDS.get(state, []))

def is_high_salary_college(name):
    low = re.search(r"\bmahavidyalaya\b", name, re.I)
    if low: return False
    return True

def get_college_website_from_cd(cd_url):
    soup = fetch(cd_url)
    if not soup: return None
    kw = ("official website","visit website","website","visit college",
          "college website","official site")
    for a in soup.find_all("a", href=True):
        href = a["href"]
        text = a.get_text(strip=True).lower()
        if "collegedunia" in href or not href.startswith("http"): continue
        if any(k in text for k in kw): return href.rstrip("/")
    for a in soup.find_all("a", href=re.compile(
            r"https?://[^/]+\.(ac\.in|edu\.in|edu)", re.I)):
        href = a["href"]
        if "collegedunia" not in href: return href.rstrip("/")
    for a in soup.find_all("a", href=re.compile(
            r"https?://(?!.*collegedunia)[^/]+\.(ac\.in|edu\.in|org\.in|edu)", re.I)):
        href = a["href"]
        if urlparse(href).path.count("/") <= 2: return href.rstrip("/")
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
    queries = [
        f'"{college_name}" official email principal contact',
        f'"{college_name}" {state} email site:ac.in OR site:edu.in',
        f'"{college_name}" principal email id',
        f'"{college_name}" HOD English email contact',
        f'{college_name} {state} college official email address',
    ]
    for q in queries:
        try:
            url = f"https://www.google.com/search?q={quote_plus(q)}&num=10"
            soup = fetch(url)
            if not soup: continue
            emails = []
            for e in EMAIL_RE.findall(soup.get_text(separator=" ")):
                e = e.lower().rstrip(".,;:)'\"")
                if e.split("@")[-1] not in SKIP_DOMAINS:
                    emails.append(e)
            if emails:
                best = best_email(emails)
                if best:
                    print(f"      [Google] {best}")
                    return best
            time.sleep(1.0)
        except Exception as ex:
            print(f"      [Google err] {ex}")
    return None

def get_email_for_college(name, website, state=""):
    if website:
        email = get_email_from_website(website)
        if email: return email
    return google_search_email(name, state)

def parse_collegedunia_pages(base_url, expected_state, pages=5):
    """Scrape multiple pages and filter by state."""
    colleges, seen_urls = [], set()
    skipped_state = skipped_salary = 0
    for page in range(1, pages + 1):
        url = base_url if page == 1 else f"{base_url}&page={page}"
        soup = fetch(url)
        if not soup: break
        found_on_page = 0
        for a in soup.find_all("a", href=re.compile(r"^/college/\d+")):
            href = "https://collegedunia.com" + a["href"]
            if href in seen_urls: continue
            name = a.get_text(strip=True)
            if not name or len(name) < 6: continue
            if re.search(
                r"[₹%#]|\bfees?\b|\bpackage\b|\branking\b|\breviews?\b"
                r"|\bbased\b|\buser\b|\d+th\b|in india|\d+\.\d+", name, re.I
            ): continue
            if not re.search(r"[a-zA-Z]{4,}", name): continue
            parent_text = a.parent.get_text(separator=" ") if a.parent else ""
            if not college_is_in_state(name + " " + parent_text, expected_state):
                skipped_state += 1
                continue
            if not is_high_salary_college(name):
                skipped_salary += 1
                continue
            seen_urls.add(href)
            colleges.append((name, href))
            found_on_page += 1
        if found_on_page == 0 and page > 1:
            break  # no more pages
        time.sleep(DELAY)
    print(f"  → {len(colleges)} from collegedunia "
          f"(skip: {skipped_state} wrong-state, {skipped_salary} low-salary)")
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

def process_college(name, state, website, seen, writer,
                    counters):
    counters['visited'] += 1
    time.sleep(DELAY)

    if not website:
        website = None  # will trigger Google fallback

    email = get_email_for_college(name, website, state)

    if not email:
        counters['no_email'] += 1
        print(f"  ✗ no email   : {name[:55]}")
        return

    email = email.lower().rstrip(".,;:)'\"")
    if email in seen:
        counters['dupes'] += 1
        return

    seen.add(email)
    writer.add(name, state, email, website or "")
    counters['saved'] += 1
    print(f"  ✓ {name[:48]!r:52s} | {email:38s}")

def main():
    sent   = load_sent()
    writer = Writer(OUTPUT, SAVE_EVERY)
    seen   = set(sent)
    counters = dict(visited=0, saved=0, no_web=0, no_email=0, dupes=0)

    print("South Indian states FIRST. Salary filter: >50k only.\n")

    # ── PHASE 1: Hardcoded seed list for South India ──────────────────────────
    print("=" * 60)
    print("PHASE 1: South Indian seed colleges (direct website scrape)")
    print("=" * 60)
    seed_states = set()
    for name, state, website in SEED_COLLEGES:
        if state not in seed_states:
            print(f"\n── {state} ──")
            seed_states.add(state)
        if not is_high_salary_college(name):
            print(f"  [skip <50k] {name[:55]}")
            continue
        process_college(name, state, website, seen, writer, counters)

    # ── PHASE 2: CollegeDunia paginated scraping ──────────────────────────────
    print("\n" + "=" * 60)
    print("PHASE 2: CollegeDunia paginated scraping (all states)")
    print("=" * 60)

    seen_cd_states = {}
    try:
        for state, base_url in COLLEGE_LIST_URLS:
            key = (state, base_url)
            if key in seen_cd_states: continue
            seen_cd_states[key] = True
            print(f"\n── {state} ({base_url.split('?')[0].split('/')[-1]}) ──")
            colleges = parse_collegedunia_pages(base_url, state, PAGES)
            for name, cd_url in colleges:
                time.sleep(DELAY)
                website = get_college_website_from_cd(cd_url)
                if not website:
                    counters['no_web'] += 1
                process_college(name, state, website, seen, writer, counters)
    except KeyboardInterrupt:
        print("\n[Stopped]")

    writer.flush()
    print(
        f"\n{'─'*60}\n"
        f"Visited:{counters['visited']}  Saved:{counters['saved']}  "
        f"NoEmail:{counters['no_email']}  Dupes:{counters['dupes']}"
    )

if __name__ == "__main__":
    main()