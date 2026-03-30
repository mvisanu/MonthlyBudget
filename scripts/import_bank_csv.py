#!/usr/bin/env python3
"""
Bank CSV -> Google Sheets importer for Personal Finance Command Center.

Usage:
    python scripts/import_bank_csv.py <csv_file> [--month 1-12] [--year 2026] [--dry-run]

Examples:
    python scripts/import_bank_csv.py ~/Downloads/303482_S9.csv
    python scripts/import_bank_csv.py ~/Downloads/303482_S9.csv --month 2 --year 2026
    python scripts/import_bank_csv.py ~/Downloads/303482_S9.csv --dry-run
"""

import csv
import io
import json
import re
import subprocess
import sys
import argparse
from collections import defaultdict
from datetime import datetime

# Fix Windows terminal emoji encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SPREADSHEET_ID = "1DEaFJvnXOM_B9GglT6sKXZzq0lCYF__EP_reMm3qu-w"

MONTH_TAB = {
    1:  "📅 Jan", 2:  "📅 Feb", 3:  "📅 Mar", 4:  "📅 Apr",
    5:  "📅 May", 6:  "📅 Jun", 7:  "📅 Jul", 8:  "📅 Aug",
    9:  "📅 Sep", 10: "📅 Oct", 11: "📅 Nov", 12: "📅 Dec",
}

# ---------------------------------------------------------------------------
# Category rules — first match wins
# ---------------------------------------------------------------------------
INCOME_RULES = [
    ("paycheck1",  ["HUMANA PAYROLL", "HUMANA"]),
    ("paycheck2",  ["BOEING", "THE BOEING"]),
    ("freelance",  ["FREELANCE", "UPWORK", "FIVERR"]),
    ("sidehustle", ["ETSY", "SHOPIFY"]),
    ("dividend",   ["DIVIDEND", "INTEREST CREDIT"]),
]

CATEGORY_RULES = [
    # NEEDS
    ("rent",        ["RENT", "LEASE", "PROPERTY MGT", "VISANU MONGSAITH", "RMTLY*"]),
    ("electric",    ["CHELCO", "CHOCTAWHATCHEE E", "GULF POWER", "FLORIDA POWER",
                     "FPL ", "DUKE ENERGY", "ELECTRIC", "EVERGY", "ENTERGY"]),
    ("water",       ["O C W & S", "OKALOOSA CO WATE", "WATER", "UTILITY"]),
    ("internet",    ["LIVEOAK FIBER", "COMCAST", "XFINITY", "AT&T INTERNET",
                     "SPECTRUM", "COX COMM", "FIBER"]),
    ("mobile",      ["ATT PAYMENT", "VERIZON", "T-MOBILE", "SPRINT", "BOOST MOBILE",
                     "METRO PCS", "AT&T"]),
    ("car_payment", ["HONDA PMT", "SANTANDER", "PNC LENDING", "TOYOTA", "FORD MOTOR",
                     "CARMAX", "CARVANA", "AUTO LOAN"]),
    ("car_ins",     ["SENTRYINS9", "DAIRYLAND", "STATE FARM RO 27 CPC",
                     "STATE FARM BILLG", "STATE FARM", "GEICO", "PROGRESSIVE INS",
                     "ALLSTATE", "NATIONWIDE INS", "FARMERS INS"]),
    ("fuel",        ["SHELL", "CHEVRON", "BP ", "EXXON", "MOBIL", "CIRCLE K",
                     "WAWA", "SPEEDWAY", "SUNOCO", "GATE PETRO", "MURPHYUSA",
                     "RACETRAC", "MARATHON"]),
    ("groceries",   ["WINN-DIXIE", "WINN DIXIE", "PUBLIX", "KROGER", "ALDI",
                     "WHOLE FOODS", "TRADER JOE", "WALMART GROCERY", "WALMART",
                     "SAMS CLUB", "SAM'S CLUB", "COSTCO", "FOOD LION",
                     "DOLLAR GENERAL MARKET", "SPROUTS"]),
    ("health_ins",  ["HEALTH INS", "BCBS", "BLUE CROSS", "CIGNA", "AETNA",
                     "UNITEDHEALTHCARE", "HUMANA HEALTH"]),
    ("medical",     ["MEDICAL", "HOSPITAL", "DR ", "DOCTOR", "PHARMACY",
                     "WALGREENS", "CVS", "RITE AID", "URGENT CARE", "CLINIC"]),
    ("debt_min",    ["CAPITAL ONE ONLINE PMT", "CREDIT ONE", "AFFIRM.COM",
                     "CONCORA CREDIT", "TOTAL CARD", "AVANT LLC", "AMZ_STORECRD",
                     "AMEX", "DISCOVER PMT", "CITI PAYMENT", "BARCLAYS"]),
    ("nsf_fee",     ["NON-SUFFICIENT FUNDS", "INSUFFICIENT FUNDS", "OVERDRAFT FEE"]),
    ("tax",         ["COMMUNITY TAX", "TAX LLC", "IRS ", "STATE TAX", "TURBOTAX"]),

    # WANTS
    ("dining",      ["MCGUIRES", "CLUBHOUSE GRILL", "DOMO CAFE", "DANNYS FRIED",
                     "DANNY'S", "RESTAURANT", "PIZZA", "MCDONALD", "CHICK-FIL",
                     "WENDY", "BURGER KING", "TACO BELL", "CHIPOTLE", "SUBWAY",
                     "STARBUCKS", "DUNKIN", "DOMINOS", "PANERA", "APPLEBEE",
                     "OLIVE GARDEN", "IHOP", "WAFFLE HOUSE", "CRACKER BARREL",
                     "SQ *", "GRUBHUB", "DOORDASH", "UBER EATS"]),
    ("streaming",   ["NETFLIX", "SPOTIFY", "DISNEY+", "HULU", "HBO MAX",
                     "APPLE TV", "PEACOCK", "PARAMOUNT", "YOUTUBE PREMIUM",
                     "GOOGLE *YOUTUBE", "AMAZON PRIME VIDEO", "TIDAL",
                     "PANDORA", "SIRIUSXM"]),
    ("subscriptions", ["ROCKET MONEY", "OPENAI", "CHATGPT", "CLAUDE.AI",
                       "CLAUDE AI", "NOTION", "LASTPASS", "ADOBE", "MICROSOFT 365",
                       "DROPBOX", "CANVA", "GRAMMARLY", "ZOOM ", "SLACK "]),
    ("shopping",    ["AMAZON", "AMZN", "TARGET", "BEST BUY", "HOME DEPOT",
                     "LOWES", "IKEA", "MARSHALLS", "TJ MAXX", "ROSS",
                     "DOLLAR TREE", "DOLLAR GENERAL", "FIVE BELOW", "EBAY"]),
    ("entertainment", ["GLF*", "GOLF", "CINEMA", "MOVIE", "TICKETMASTER",
                       "STUBHUB", "AMC ", "REGAL ", "BOWLING", "ARCADE",
                       "MUSEUM", "CONCERT"]),
    ("gym",         ["PLANET FITNESS", "ANYTIME FITNESS", "LA FITNESS", "YMCA",
                     "GYM", "FITNESS"]),
    ("carwash",     ["CAR WASH", "TAKE 5", "MISTER CAR", "CLEAN FREAK"]),
    ("beauty",      ["SALON", "BARBER", "NAIL", "SPA ", "BEAUTY"]),
    ("travel",      ["HOTEL", "AIRBNB", "VRBO", "EXPEDIA", "KAYAK",
                     "DELTA ", "AMERICAN AIR", "SOUTHWEST", "UNITED AIR",
                     "SPIRIT AIR", "CARNIVAL", "ROYAL CARIBBEAN"]),
    ("cash_atm",    ["ATM WITHDRAWAL", "CASH WITHDRAWAL"]),

    # SAVINGS & DEBT
    ("retirement",   ["TO YOUR LOAN 401", "LOAN 401", "FIDELITY", "VANGUARD",
                      "SCHWAB", "TIAA", "401K", "IRA CONTRIBUTION"]),
    ("investment",   ["ROBINHOOD", "STASH", "ACORNS", "WEBULL", "TD AMERITRADE",
                      "ETRADE", "M1 FINANCE", "BETTERMENT", "WEALTHFRONT"]),
    ("savings_xfer", ["TO YOUR SHARE", "TO YOUR SAVINGS", "ONLINE BANKING TRANSFER"]),
    ("extra_debt",   ["TO YOUR LOAN 3", "TO YOUR LOAN 4", "LOAN PAYMENT"]),
]

# Map category key -> sheet cell (column B = Spent $) in new monthly tab layout
SHEET_ROWS = {
    # Income — rows 6-9 col B
    "paycheck1":     "B6",  "paycheck2":    "B7",
    "freelance":     "B8",  "sidehustle":   "B8",  "dividend": "B9",
    # Fixed expenses — rows 14-23 col B
    "rent":          "B14", "car_payment":  "B15", "car_ins":   "B16",
    "health_ins":    "B17", "internet":     "B18", "mobile":    "B19",
    "electric":      "B20", "water":        "B21", "debt_min":  "B22",
    "retirement":    "B23",
    # Essentials — rows 28-30 col B
    "groceries":     "B28", "fuel":         "B29", "medical":   "B30",
    # Discretionary — rows 35-46 col B
    "dining":        "B35", "entertainment":"B36", "shopping":  "B37",
    "subscriptions": "B38", "streaming":    "B39", "carwash":   "B40",
    "gym":           "B41", "beauty":       "B42", "travel":    "B43",
    "cash_atm":      "B44", "nsf_fee":      "B45", "tax":       "B46",
    # Savings & extra debt — rows 51-53 col B
    "extra_debt":    "B51", "savings_xfer": "B52", "investment":"B53",
}

CATEGORY_LABELS = {
    "paycheck1": "Paycheck (Humana)", "paycheck2": "Paycheck (Boeing)",
    "freelance": "Freelance", "sidehustle": "Side Hustle",
    "dividend": "Dividend/Interest",
    "rent": "Rent/Housing", "electric": "Electricity",
    "water": "Water/Utilities", "internet": "Internet",
    "mobile": "Mobile Phone", "car_payment": "Car Payment",
    "car_ins": "Car Insurance", "fuel": "Fuel",
    "groceries": "Groceries", "health_ins": "Health Insurance",
    "medical": "Medical", "debt_min": "Min Debt Payments",
    "tax": "Taxes", "nsf_fee": "NSF/Bank Fees",
    "dining": "Dining Out", "entertainment": "Entertainment",
    "streaming": "Streaming", "shopping": "Shopping",
    "gym": "Gym", "beauty": "Beauty/Personal",
    "travel": "Travel", "carwash": "Car Wash",
    "cash_atm": "Cash / ATM", "subscriptions": "Subscriptions",
    "retirement": "Retirement/401k", "investment": "Investments",
    "extra_debt": "Extra Debt Payment", "savings_xfer": "Savings Transfer",
    "uncategorized": "Uncategorized",
}

def _section(key):
    needs   = {"rent","electric","water","internet","mobile","car_payment","car_ins",
               "fuel","groceries","health_ins","medical","debt_min","nsf_fee","tax"}
    wants   = {"dining","streaming","subscriptions","shopping","entertainment",
               "gym","carwash","beauty","travel","cash_atm"}
    savings = {"retirement","investment","savings_xfer","extra_debt"}
    if key in needs:   return "Needs"
    if key in wants:   return "Wants"
    if key in savings: return "Savings/Debt"
    return "Other"

def categorize(description, ext, amount):
    desc_upper = (description or "").upper()
    ext_upper  = (ext or "").upper()

    if "OVERDRAFT TRANSFER" in ext_upper:
        return ("skip", "overdraft")
    if "Online Banking Deposit" in ext or "FROM YOUR SHARE" in desc_upper:
        return ("skip", "internal_transfer")

    # Cash withdrawal = blank description + Share Withdrawal
    if not description.strip() and "Share Withdrawal" in ext:
        return ("Wants", "cash_atm")

    if amount > 0 and "CREDIT" in ext_upper:
        for key, keywords in INCOME_RULES:
            if any(k in desc_upper for k in keywords):
                return ("Income", key)
        return ("Income", "other_income")

    for key, keywords in CATEGORY_RULES:
        if any(k.upper() in desc_upper for k in keywords):
            return (_section(key), key)

    return ("Other", "uncategorized")

def parse_csv(filepath):
    transactions = []
    with open(filepath, newline="", encoding="utf-8-sig") as f:
        raw = f.read().lstrip("\n\r")
    reader = csv.DictReader(raw.splitlines())
    for row in reader:
        date_str = row.get("Date", "").strip()
        if not date_str:
            continue
        try:
            date = datetime.strptime(date_str, "%m/%d/%Y")
        except ValueError:
            continue
        amount_str = row.get("Amount", "0").strip().replace(",", "")
        try:
            amount = float(amount_str)
        except ValueError:
            amount = 0.0
        transactions.append({
            "date":   date,
            "desc":   row.get("Description", "").strip(),
            "ext":    row.get("Ext", "").strip(),
            "amount": amount,
        })
    return transactions

def gws(params, payload):
    """Run a gws values batchUpdate call. Returns (ok, stderr)."""
    result = subprocess.run(
        ["gws.cmd", "sheets", "spreadsheets", "values", "batchUpdate",
         "--params", json.dumps(params),
         "--json",   json.dumps(payload, ensure_ascii=False)],
        capture_output=True, text=True, encoding="utf-8"
    )
    return result.returncode == 0, result.stderr

def clear_month_transactions(month, year):
    """Clear existing transaction rows for this month from the Transactions tab."""
    # We'll read existing data and rewrite without this month's rows
    # For simplicity, we clear all and rewrite — handled in write_transactions
    pass

def write_transactions(all_transactions_by_month):
    """
    Write all transaction rows to the Transactions tab.
    Clears existing data first, then rewrites everything.
    Format: Date | Payee | Amount | Type | Category | Section | Month
    """
    header = [["Date", "Payee / Description", "Amount ($)", "Txn Type",
               "Category", "Section", "Month"]]

    rows = []
    for (month, year), txns in sorted(all_transactions_by_month.items()):
        month_label = datetime(year, month, 1).strftime("%B %Y")
        for t in sorted(txns, key=lambda x: x["date"]):
            section, key = categorize(t["desc"], t["ext"], t["amount"])
            if section == "skip":
                continue
            label = CATEGORY_LABELS.get(key, key)
            raw = t["desc"] if t["desc"] else f"[Cash/ATM - {t['ext']}]"
            # Strip bank ref codes like "0PRR3J|A953472715 " at the start
            clean = re.sub(r'^[A-Z0-9]{4,8}\|[A-Z0-9]{6,12}\s+', '', raw).strip()
            # Remove card numbers (************1234)
            clean = re.sub(r'\*{4,}\d{4}', '', clean)
            # Remove Windows cmd-special chars that break subprocess
            for ch in ('|', '&', '<', '>', '^', '%', '!', '#', '"'):
                clean = clean.replace(ch, ' ')
            # Collapse multiple spaces and trim to 60 chars
            payee = re.sub(r' {2,}', ' ', clean).strip()[:60]
            rows.append([
                t["date"].strftime("%m/%d/%Y"),
                payee,
                round(t["amount"], 2),
                t["ext"],
                label,
                section,
                month_label,
            ])

    params = {"spreadsheetId": SPREADSHEET_ID}

    # Clear existing data first
    subprocess.run(
        ["gws.cmd", "sheets", "spreadsheets", "values", "clear",
         "--params", json.dumps({
             "spreadsheetId": SPREADSHEET_ID,
             "range": "'📋 Transactions'!A1:G2000"
         })],
        capture_output=True, text=True, encoding="utf-8"
    )

    # Write header + data in chunks of 25 rows to stay under Windows cmd limit
    all_rows = header + rows
    chunk_size = 8
    start_row = 1
    for i in range(0, len(all_rows), chunk_size):
        chunk = all_rows[i:i + chunk_size]
        payload = {
            "valueInputOption": "USER_ENTERED",
            "data": [{"range": f"'📋 Transactions'!A{start_row}",
                      "values": chunk}]
        }
        ok, err = gws(params, payload)
        if not ok:
            print(f"    Transactions write error (chunk {i}): {err[:200]}")
            return False, len(rows)
        start_row += len(chunk)

    return True, len(rows)

def format_transactions_tab(sheet_id=2114397835):
    """Apply formatting to the Transactions tab."""
    navy  = {"red": 0.118, "green": 0.227, "blue": 0.373}
    white = {"red": 1, "green": 1, "blue": 1}
    gray1 = {"red": 0.976, "green": 0.980, "blue": 0.984}  # #F9FAFB
    gray2 = {"red": 1.0,   "green": 1.0,   "blue": 1.0}

    requests = [
        # Header row — navy bg, white bold text
        {"repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {
                "backgroundColor": navy,
                "textFormat": {"bold": True, "fontSize": 11,
                               "foregroundColor": white},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }},
        # Data rows alternating
        {"repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2000},
            "cell": {"userEnteredFormat": {
                "textFormat": {"fontSize": 10},
                "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat(textFormat,verticalAlignment)"
        }},
        # Freeze header row
        {"updateSheetProperties": {
            "properties": {"sheetId": sheet_id,
                           "gridProperties": {"frozenRowCount": 1}},
            "fields": "gridProperties.frozenRowCount"
        }},
        # Column widths: Date=90, Payee=350, Amount=100, Type=160, Category=160, Section=100, Month=110
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 95}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 1, "endIndex": 2},
            "properties": {"pixelSize": 360}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 2, "endIndex": 3},
            "properties": {"pixelSize": 105}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 3, "endIndex": 4},
            "properties": {"pixelSize": 165}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 4, "endIndex": 5},
            "properties": {"pixelSize": 165}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 5, "endIndex": 6},
            "properties": {"pixelSize": 105}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 6, "endIndex": 7},
            "properties": {"pixelSize": 115}, "fields": "pixelSize"
        }},
        # Row height
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 28}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": 1, "endIndex": 2000},
            "properties": {"pixelSize": 22}, "fields": "pixelSize"
        }},
        # Light background for all data rows
        {"repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2000,
                      "startColumnIndex": 0, "endColumnIndex": 7},
            "cell": {"userEnteredFormat": {"backgroundColor": gray1}},
            "fields": "userEnteredFormat.backgroundColor"
        }},
        # Amount column: right-align
        {"repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2000,
                      "startColumnIndex": 2, "endColumnIndex": 3},
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00;($#,##0.00);\"-\""},
                "horizontalAlignment": "RIGHT",
            }},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
        }},
        # Borders around entire table
        {"updateBorders": {
            "range": {"sheetId": sheet_id,
                      "startRowIndex": 0, "endRowIndex": 2000,
                      "startColumnIndex": 0, "endColumnIndex": 7},
            "innerVertical": {"style": "SOLID", "color": {"red": 0.82, "green": 0.82, "blue": 0.82}},
            "innerHorizontal": {"style": "SOLID", "color": {"red": 0.88, "green": 0.88, "blue": 0.88}},
        }},
    ]

    params = {"spreadsheetId": SPREADSHEET_ID}
    result = subprocess.run(
        ["gws.cmd", "sheets", "spreadsheets", "batchUpdate",
         "--params", json.dumps(params),
         "--json",   json.dumps({"requests": requests}, ensure_ascii=False)],
        capture_output=True, text=True, encoding="utf-8"
    )
    return result.returncode == 0

def format_monthly_tab(sheet_id, month_name):
    """Apply clean formatting to a monthly budget tab."""
    navy    = {"red": 0.118, "green": 0.227, "blue": 0.373}
    blue_md = {"red": 0.145, "green": 0.388, "blue": 0.922}
    blue_lt = {"red": 0.859, "green": 0.918, "blue": 0.996}
    white   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
    navy_t  = {"red": 0.118, "green": 0.227, "blue": 0.373}
    gray_lt = {"red": 0.953, "green": 0.957, "blue": 0.965}
    yellow  = {"red": 0.996, "green": 0.988, "blue": 0.910}
    blue_in = {"red": 0.114, "green": 0.306, "blue": 0.847}
    green_t = {"red": 0.082, "green": 0.502, "blue": 0.239}
    alt1    = {"red": 0.976, "green": 0.980, "blue": 0.984}
    pos_bg  = {"red": 0.820, "green": 0.980, "blue": 0.898}
    neg_bg  = {"red": 0.996, "green": 0.886, "blue": 0.886}

    def cell_range(r1, c1, r2, c2):
        return {"sheetId": sheet_id,
                "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    def header_fmt(bg, text_color, size=11, bold=True):
        return {
            "backgroundColor": bg,
            "textFormat": {"bold": bold, "fontSize": size,
                           "foregroundColor": text_color},
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE",
        }

    def data_fmt(text_color=None, bold=False, bg=None, align="LEFT"):
        fmt = {"textFormat": {"bold": bold, "fontSize": 10,
                               "foregroundColor": text_color or {"red":0,"green":0,"blue":0}},
               "horizontalAlignment": align,
               "verticalAlignment": "MIDDLE"}
        if bg:
            fmt["backgroundColor"] = bg
        return fmt

    # Section header rows (0-indexed): title=0, subtitle=1, period=3
    # Income header=5, income cols=6, income data=7-14, income total=15
    # Needs header=17, needs cols=18, needs data=19-33, needs total=33
    # Wants header=35, wants cols=36, wants data=37-49, wants total=49
    # Savings header=51, savings cols=52, savings data=53-61, savings total=61
    # Totals header=63, totals data=64-70

    section_headers = [0, 1, 5, 17, 35, 51, 63]
    navy_rows  = [0, 5, 17, 35, 51, 63]
    blue_rows  = [1]
    col_header_rows = [6, 18, 36, 52, 64]
    total_rows = [15, 33, 49, 61, 68, 69, 70]
    income_data  = list(range(7, 15))
    needs_data   = list(range(19, 33))
    wants_data   = list(range(37, 49))
    savings_data = list(range(53, 61))

    requests = [
        # Freeze first 7 rows + column A
        {"updateSheetProperties": {
            "properties": {"sheetId": sheet_id,
                           "gridProperties": {"frozenRowCount": 7, "frozenColumnCount": 1}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
        }},
        # Row 1 — title (navy)
        {"repeatCell": {
            "range": cell_range(0, 0, 1, 6),
            "cell": {"userEnteredFormat": header_fmt(navy, white, 14)},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }},
        # Row 2 — subtitle (medium blue)
        {"repeatCell": {
            "range": cell_range(1, 0, 2, 6),
            "cell": {"userEnteredFormat": header_fmt(blue_md, white, 11)},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }},
        # Section header rows (navy)
        *[{"repeatCell": {
            "range": cell_range(r, 0, r+1, 6),
            "cell": {"userEnteredFormat": header_fmt(navy, white, 11)},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }} for r in navy_rows[2:]],  # skip row 0 already done
        # Column header rows (light blue)
        *[{"repeatCell": {
            "range": cell_range(r, 0, r+1, 6),
            "cell": {"userEnteredFormat": {
                "backgroundColor": blue_lt,
                "textFormat": {"bold": True, "fontSize": 10,
                               "foregroundColor": navy_t},
                "horizontalAlignment": "CENTER",
                "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
        }} for r in col_header_rows],
        # Total rows (light gray, bold)
        *[{"repeatCell": {
            "range": cell_range(r, 0, r+1, 6),
            "cell": {"userEnteredFormat": {
                "backgroundColor": gray_lt,
                "textFormat": {"bold": True, "fontSize": 10},
                "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)"
        }} for r in total_rows],
        # Alternating data rows — income
        *[{"repeatCell": {
            "range": cell_range(r, 0, r+1, 6),
            "cell": {"userEnteredFormat": {
                "backgroundColor": alt1 if i % 2 == 0 else white,
                "textFormat": {"fontSize": 10},
                "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)"
        }} for i, r in enumerate(income_data + needs_data + wants_data + savings_data)],
        # Column widths: A=200, B=120, C=110, D=110, E=60, F=80
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 200}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 1, "endIndex": 2},
            "properties": {"pixelSize": 110}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 2, "endIndex": 4},
            "properties": {"pixelSize": 115}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 4, "endIndex": 5},
            "properties": {"pixelSize": 65}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 5, "endIndex": 6},
            "properties": {"pixelSize": 85}, "fields": "pixelSize"
        }},
        # Row heights: title rows 30px, section headers 26px, data rows 22px
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": 0, "endIndex": 2},
            "properties": {"pixelSize": 30}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": 2, "endIndex": 72},
            "properties": {"pixelSize": 22}, "fields": "pixelSize"
        }},
        # Currency format on C, D columns (Expected/Actual)
        {"repeatCell": {
            "range": cell_range(7, 2, 72, 4),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00;($#,##0.00);\"-\""},
                "horizontalAlignment": "RIGHT",
            }},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
        }},
        # Variance column E: number format + right align
        {"repeatCell": {
            "range": cell_range(7, 4, 72, 5),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00;($#,##0.00);\"-\""},
                "horizontalAlignment": "RIGHT",
            }},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
        }},
        # Progress % column F
        {"repeatCell": {
            "range": cell_range(7, 5, 72, 6),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "0.0%"},
                "horizontalAlignment": "CENTER",
            }},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
        }},
        # Net Cash Flow row (row 70, index 69) — bold + larger
        {"repeatCell": {
            "range": cell_range(69, 0, 70, 6),
            "cell": {"userEnteredFormat": {
                "backgroundColor": navy,
                "textFormat": {"bold": True, "fontSize": 12, "foregroundColor": white},
                "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)"
        }},
        # Conditional formatting: variance positive = green bg, negative = red bg
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [cell_range(7, 4, 72, 5)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_GREATER", "values": [{"userEnteredValue": "0"}]},
                    "format": {"backgroundColor": pos_bg}
                }
            }, "index": 0
        }},
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [cell_range(7, 4, 72, 5)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                    "format": {"backgroundColor": neg_bg}
                }
            }, "index": 1
        }},
        # Inner borders on data area
        {"updateBorders": {
            "range": cell_range(6, 0, 72, 6),
            "innerHorizontal": {"style": "SOLID",
                                "color": {"red": 0.85, "green": 0.85, "blue": 0.85}},
            "innerVertical":   {"style": "SOLID",
                                "color": {"red": 0.85, "green": 0.85, "blue": 0.85}},
        }},
    ]

    params = {"spreadsheetId": SPREADSHEET_ID}
    # Split into chunks of 10 requests to avoid command-line length limits
    chunk_size = 10
    for i in range(0, len(requests), chunk_size):
        chunk = requests[i:i + chunk_size]
        result = subprocess.run(
            ["gws.cmd", "sheets", "spreadsheets", "batchUpdate",
             "--params", json.dumps(params),
             "--json",   json.dumps({"requests": chunk}, ensure_ascii=False)],
            capture_output=True, text=True, encoding="utf-8"
        )
        if result.returncode != 0:
            return False, result.stderr
    return True, ""

# Sheet IDs for all monthly tabs
MONTHLY_SHEET_IDS = {
    1: 122166217,  2: 636285504,  3: 1742587835, 4: 278726061,
    5: 1054479217, 6: 1278089603, 7: 542569607,  8: 555749728,
    9: 1955324181, 10: 679307036, 11: 1270203974, 12: 1885490765,
}

def aggregate(transactions, month, year):
    income  = defaultdict(float)
    expense = defaultdict(float)
    unknown = []
    for t in transactions:
        if t["date"].month != month or t["date"].year != year:
            continue
        section, key = categorize(t["desc"], t["ext"], t["amount"])
        if section == "skip":
            continue
        elif section == "Income":
            income[key] += t["amount"]
        elif key != "uncategorized":
            expense[key] += abs(t["amount"])
        else:
            unknown.append(t)
    return income, expense, unknown

def build_value_updates(income, expense, tab_name):
    updates = []
    def add(cell, value):
        updates.append({
            "range": f"'{tab_name}'!{cell}",
            "values": [[round(value, 2)]]
        })
    for key, val in income.items():
        if key in SHEET_ROWS:
            add(SHEET_ROWS[key], val)
    for key, val in expense.items():
        if key in SHEET_ROWS:
            add(SHEET_ROWS[key], val)
    return updates

def print_summary(income, expense, unknown, month, year):
    print(f"\n{'='*54}")
    print(f"  {datetime(year, month, 1).strftime('%B %Y')}")
    print(f"{'='*54}")
    total_in = sum(income.values())
    total_out = sum(expense.values())
    print(f"  Income:   ${total_in:>10,.2f}")
    print(f"  Expenses: ${total_out:>10,.2f}")
    print(f"  Net:      ${total_in - total_out:>10,.2f}")
    if unknown:
        print(f"  Uncategorized ({len(unknown)}):")
        for t in unknown[:10]:
            print(f"    {t['date'].strftime('%m/%d')}  ${t['amount']:>9,.2f}  {t['desc'][:50]}")

def main():
    parser = argparse.ArgumentParser(description="Import bank CSV into Google Sheets budget.")
    parser.add_argument("csv_file", help="Path to bank CSV export")
    parser.add_argument("--month", type=int, default=None)
    parser.add_argument("--year",  type=int, default=None)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--skip-format", action="store_true",
                        help="Skip re-applying tab formatting (faster re-runs)")
    args = parser.parse_args()

    print(f"Reading {args.csv_file}...")
    transactions = parse_csv(args.csv_file)
    if not transactions:
        print("No transactions found.")
        sys.exit(1)
    print(f"  Loaded {len(transactions)} transactions.")

    from collections import Counter
    month_counts = Counter((t["date"].month, t["date"].year) for t in transactions)
    if args.month or args.year:
        month_counts = {(m, y): c for (m, y), c in month_counts.items()
                        if (args.month is None or m == args.month)
                        and (args.year  is None or y == args.year)}

    periods = sorted(month_counts.keys())
    print(f"  Months found: " +
          ", ".join(datetime(y, m, 1).strftime("%b %Y") for m, y in periods))

    if args.dry_run:
        for (month, year) in periods:
            income, expense, unknown = aggregate(transactions, month, year)
            print_summary(income, expense, unknown, month, year)
        print("\n[DRY RUN] No changes written.")
        return

    # --- Write category totals to monthly tabs ---
    for (month, year) in periods:
        tab_name = MONTH_TAB.get(month)
        income, expense, unknown = aggregate(transactions, month, year)
        print_summary(income, expense, unknown, month, year)
        updates = build_value_updates(income, expense, tab_name)
        if updates:
            params = {"spreadsheetId": SPREADSHEET_ID}
            ok, err = gws(params, {
                "valueInputOption": "USER_ENTERED",
                "data": updates
            })
            print(f"  {'OK' if ok else 'ERROR'} — wrote {len(updates)} cells to '{tab_name}'")
            if not ok:
                print(f"  {err[:200]}")

    # --- Write all transactions to Transactions tab ---
    print("\n  Updating Transactions tab...")
    txns_by_month = {(m, y): [] for m, y in periods}
    for t in transactions:
        key = (t["date"].month, t["date"].year)
        if key in txns_by_month:
            txns_by_month[key].append(t)

    ok, count = write_transactions(txns_by_month)
    print(f"  {'OK' if ok else 'ERROR'} — wrote {count} transaction rows")

    # --- Apply formatting ---
    if not args.skip_format:
        print("\n  Applying formatting...")
        # Format Transactions tab
        ok = format_transactions_tab()
        print(f"  Transactions tab: {'OK' if ok else 'ERROR'}")

        # Monthly tab formatting is handled by setup_monthly_tabs.py
        print("  (monthly tab formatting managed by setup_monthly_tabs.py)")
    else:
        print("  (formatting skipped)")

    print(f"\nDone. Open your sheet:")
    print(f"  https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")

if __name__ == "__main__":
    main()
