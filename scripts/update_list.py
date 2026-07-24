"""Weekly FEHD list sync for seating.xlsx.

Run by GitHub Actions on a schedule. Steps, in order:
1. If the live FEHD list differs from the sheet, back up the current
   sheet as 'seating prior.xlsx' BEFORE changing anything.
2. Append rows for restaurants newly added by FEHD (seating left blank).
3. Report removed licences in the commit summary (rows are left in place;
   the app ignores them automatically).
Writes summary.txt for the commit message. Exits quietly with no file
changes when the list is unchanged.
"""
import shutil
import sys

import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

SHEET = "seating.xlsx"
BACKUP = "seating prior.xlsx"
URL = "https://www.fehd.gov.hk/english/licensing/dog_restaurants/getData.php"
REFERER = "https://www.fehd.gov.hk/english/licensing/dog_restaurants/dog_restaurants_list.html"

resp = requests.get(URL, headers={
    "Accept": "application/json",
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)",
    "Referer": REFERER,
}, timeout=60)
resp.raise_for_status()
data = resp.json()
records = data if isinstance(data, list) else next(iter(data.values()))
live = {}
for x in records:
    lic = str(x.get("licence", "")).strip()
    if lic:
        live[lic] = x
if len(live) < 500:
    sys.exit(f"FEHD returned only {len(live)} records - refusing to proceed")

wb = load_workbook(SHEET)
ws = wb["Seating"]
existing = set()
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0] is not None:
        existing.add(str(row[0]).strip())

added = [live[l] for l in live if l not in existing]
removed = sorted(existing - set(live))

if not added and not removed:
    print("FEHD list unchanged - nothing to do")
    sys.exit(0)

if added:
    # 1. Backup BEFORE modifying anything.
    shutil.copyfile(SHEET, BACKUP)

    # 2. Append the new restaurants.
    yellow = PatternFill("solid", fgColor="FFF3C4")
    added.sort(key=lambda x: (x.get("district_en", ""), x.get("shop_sign_en", "")))
    for x in added:
        r = ws.max_row + 1
        values = [str(x.get("licence", "")).strip(), x.get("shop_sign_en", ""),
                  x.get("shop_sign_tc", ""), x.get("district_en", ""),
                  x.get("address_en", ""), ""]
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name="Arial", size=10)
            if c == 1:
                cell.number_format = "@"
            if c == 6:
                cell.fill = yellow
                cell.font = Font(name="Arial", size=10, bold=True)

    if ws.data_validations.dataValidation:
        ws.data_validations.dataValidation[0].sqref = f"F2:F{ws.max_row}"
    ws.auto_filter.ref = f"A1:F{ws.max_row}"
    wb.save(SHEET)

# 3. Summary for the commit message.
lines = [f"FEHD sync: {len(added)} added, {len(removed)} delisted", ""]
for x in added:
    lines.append(f"+ [{str(x.get('licence','')).strip()}] {x.get('shop_sign_en','')} ({x.get('district_en','')})")
for l in removed:
    lines.append(f"- [{l}] delisted by FEHD (row kept; app hides it automatically)")
with open("summary.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(lines) + "\n")
print("\n".join(lines))
