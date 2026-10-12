"""
Step 1 — Run this ONCE to build sent_emails.txt from all your old Excel files.
Put all your old contacts_list_*.xlsx files in the same folder as this script.
"""
import glob
import openpyxl

sent = set()

# Find all xlsx files in current folder
files = glob.glob("*.xlsx") + glob.glob("contacts_list_*.xlsx") + glob.glob("facultytick_contacts_*.xlsx")
files = list(set(files))  # dedup

if not files:
    print("No xlsx files found in current folder.")
else:
    for fname in sorted(files):
        try:
            wb = openpyxl.load_workbook(fname)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):  # skip header
                email = str(row[1]).strip().lower() if row[1] else ""
                if "@" in email:
                    sent.add(email)
            print(f"  ✓ {fname} — loaded")
        except Exception as e:
            print(f"  ✗ {fname} — {e}")

with open("sent_emails.txt", "w") as f:
    for email in sorted(sent):
        f.write(email + "\n")

print(f"\nDone — {len(sent)} unique emails saved to sent_emails.txt")
print("Now run scrape_facultyplus.py to get only NEW contacts.")