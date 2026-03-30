#!/usr/bin/env python3
"""
Rewrites all 12 monthly tabs with a clean spending-tracker layout.
Run once to set up. Re-running is safe (clears and rewrites).

Usage:
    python scripts/setup_monthly_tabs.py
"""

import io, json, subprocess, sys, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SPREADSHEET_ID = "1DEaFJvnXOM_B9GglT6sKXZzq0lCYF__EP_reMm3qu-w"

MONTHS = [
    (1,  "📅 Jan", "JANUARY",   122166217),
    (2,  "📅 Feb", "FEBRUARY",  636285504),
    (3,  "📅 Mar", "MARCH",     1742587835),
    (4,  "📅 Apr", "APRIL",     278726061),
    (5,  "📅 May", "MAY",       1054479217),
    (6,  "📅 Jun", "JUNE",      1278089603),
    (7,  "📅 Jul", "JULY",      542569607),
    (8,  "📅 Aug", "AUGUST",    555749728),
    (9,  "📅 Sep", "SEPTEMBER", 1955324181),
    (10, "📅 Oct", "OCTOBER",   679307036),
    (11, "📅 Nov", "NOVEMBER",  1270203974),
    (12, "📅 Dec", "DECEMBER",  1885490765),
]

# ── Layout (row numbers are 1-based for display, 0-based in API) ─────────────
#
#  Row  1  : Title — "JANUARY 2026 — Spending Tracker"
#  Row  2  : Snapshot formula — Income / Spent / Left
#  Row  3  : blank
#  Row  4  : INCOME  (section header)
#  Row  5  : col headers
#  Row  6  : Paycheck 1 (Humana)     ← B6  import target
#  Row  7  : Paycheck 2 (Boeing)     ← B7
#  Row  8  : Freelance               ← B8
#  Row  9  : Other Income            ← B9
#  Row 10  : TOTAL INCOME            ← B10 = SUM(B6:B9)
#  Row 11  : blank
#  Row 12  : FIXED EXPENSES  (navy)
#  Row 13  : col headers
#  Row 14  : Rent / Housing          ← B14
#  Row 15  : Car Payment(s)          ← B15
#  Row 16  : Car Insurance           ← B16
#  Row 17  : Health Insurance        ← B17
#  Row 18  : Internet                ← B18
#  Row 19  : Mobile Phone            ← B19
#  Row 20  : Electricity             ← B20
#  Row 21  : Water / Utilities       ← B21
#  Row 22  : Min Debt Payments       ← B22
#  Row 23  : Retirement / 401k       ← B23
#  Row 24  : SUBTOTAL FIXED          ← B24 = SUM(B14:B23)
#  Row 25  : blank
#  Row 26  : ESSENTIALS  (navy)
#  Row 27  : col headers
#  Row 28  : Groceries               ← B28
#  Row 29  : Fuel / Gas              ← B29
#  Row 30  : Medical / Pharmacy      ← B30
#  Row 31  : SUBTOTAL ESSENTIALS     ← B31 = SUM(B28:B30)
#  Row 32  : blank
#  Row 33  : DISCRETIONARY  (dark red/orange)
#  Row 34  : col headers
#  Row 35  : Dining Out              ← B35
#  Row 36  : Entertainment / Golf    ← B36
#  Row 37  : Shopping                ← B37
#  Row 38  : Subscriptions           ← B38
#  Row 39  : Streaming               ← B39
#  Row 40  : Car Wash                ← B40
#  Row 41  : Gym / Fitness           ← B41
#  Row 42  : Beauty / Personal       ← B42
#  Row 43  : Travel                  ← B43
#  Row 44  : Cash / ATM Withdrawals  ← B44
#  Row 45  : Bank Fees (NSF)         ← B45
#  Row 46  : Taxes / Tax Services    ← B46
#  Row 47  : SUBTOTAL DISCRETIONARY  ← B47 = SUM(B35:B46)
#  Row 48  : blank
#  Row 49  : SAVINGS & EXTRA DEBT  (green-navy)
#  Row 50  : col headers
#  Row 51  : Extra Debt Payments     ← B51
#  Row 52  : Savings Transfers       ← B52
#  Row 53  : Investments             ← B53
#  Row 54  : SUBTOTAL SAVINGS        ← B54 = SUM(B51:B53)
#  Row 55  : blank
#  Row 56  : SUMMARY  (navy)
#  Row 57  : Total Income            B57=B10
#  Row 58  : Fixed Expenses          B58=B24
#  Row 59  : Essentials              B59=B31
#  Row 60  : Discretionary           B60=B47
#  Row 61  : Savings/Extra Debt      B61=B54
#  Row 62  : TOTAL SPENT             B62=B58+B59+B60+B61
#  Row 63  : blank
#  Row 64  : MONEY LEFT OVER         B64=B57-B62   ← THE BIG NUMBER
#  Row 65  : Savings Rate            B65=IFERROR(B54/B10,0)

def tab_values(month_name, year=2026):
    T = f"{month_name} {year}"
    pct = "=IFERROR(B{row}/$B$10,0)"

    rows = [
        # Row 1 — title
        [f"{month_name} {year} — Spending Tracker", "", "", "", ""],
        # Row 2 — snapshot (split across cells to avoid nested quotes in cmd)
        ["Income:", "=B10", "Spent:", "=B62", "Left Over:", "=B64"],
        # Row 3 — blank
        ["", "", "", "", ""],
        # Row 4 — INCOME header
        ["INCOME", "", "", "", ""],
        # Row 5 — col headers
        ["Source", "Amount ($)", "Budget ($)", "vs Budget", "Pct of Income"],
        # Row 6-9 — income rows
        ["Paycheck 1 (Humana)",  "", "", "=IF(C6=\"\",\"\",C6-B6)",  "=IFERROR(B6/$B$10,0)"],
        ["Paycheck 2 (Boeing)",  "", "", "=IF(C7=\"\",\"\",C7-B7)",  "=IFERROR(B7/$B$10,0)"],
        ["Freelance / Side Hustle", "", "", "=IF(C8=\"\",\"\",C8-B8)",  "=IFERROR(B8/$B$10,0)"],
        ["Other Income",         "", "", "=IF(C9=\"\",\"\",C9-B9)",  "=IFERROR(B9/$B$10,0)"],
        # Row 10 — total income
        ["TOTAL INCOME", "=SUM(B6:B9)", "", "", ""],
        # Row 11 — blank
        ["", "", "", "", ""],
        # Row 12 — FIXED EXPENSES header
        ["FIXED EXPENSES", "", "", "", ""],
        # Row 13 — col headers
        ["Category", "Spent ($)", "Budget ($)", "vs Budget", "Pct of Income"],
        # Rows 14-23 — fixed categories
        ["Rent / Housing",        "", "", "=IF(C14=\"\",\"\",C14-B14)", "=IFERROR(B14/$B$10,0)"],
        ["Car Payment(s)",        "", "", "=IF(C15=\"\",\"\",C15-B15)", "=IFERROR(B15/$B$10,0)"],
        ["Car Insurance",         "", "", "=IF(C16=\"\",\"\",C16-B16)", "=IFERROR(B16/$B$10,0)"],
        ["Health Insurance",      "", "", "=IF(C17=\"\",\"\",C17-B17)", "=IFERROR(B17/$B$10,0)"],
        ["Internet",              "", "", "=IF(C18=\"\",\"\",C18-B18)", "=IFERROR(B18/$B$10,0)"],
        ["Mobile Phone",          "", "", "=IF(C19=\"\",\"\",C19-B19)", "=IFERROR(B19/$B$10,0)"],
        ["Electricity",           "", "", "=IF(C20=\"\",\"\",C20-B20)", "=IFERROR(B20/$B$10,0)"],
        ["Water / Utilities",     "", "", "=IF(C21=\"\",\"\",C21-B21)", "=IFERROR(B21/$B$10,0)"],
        ["Min Debt Payments",     "", "", "=IF(C22=\"\",\"\",C22-B22)", "=IFERROR(B22/$B$10,0)"],
        ["Retirement / 401k",     "", "", "=IF(C23=\"\",\"\",C23-B23)", "=IFERROR(B23/$B$10,0)"],
        # Row 24 — subtotal fixed
        ["SUBTOTAL — FIXED", "=SUM(B14:B23)", "", "", "=IFERROR(B24/$B$10,0)"],
        # Row 25 — blank
        ["", "", "", "", ""],
        # Row 26 — ESSENTIALS header
        ["ESSENTIALS", "", "", "", ""],
        # Row 27 — col headers
        ["Category", "Spent ($)", "Budget ($)", "vs Budget", "Pct of Income"],
        # Rows 28-30
        ["Groceries",            "", "", "=IF(C28=\"\",\"\",C28-B28)", "=IFERROR(B28/$B$10,0)"],
        ["Fuel / Gas",           "", "", "=IF(C29=\"\",\"\",C29-B29)", "=IFERROR(B29/$B$10,0)"],
        ["Medical / Pharmacy",   "", "", "=IF(C30=\"\",\"\",C30-B30)", "=IFERROR(B30/$B$10,0)"],
        # Row 31 — subtotal
        ["SUBTOTAL — ESSENTIALS", "=SUM(B28:B30)", "", "", "=IFERROR(B31/$B$10,0)"],
        # Row 32 — blank
        ["", "", "", "", ""],
        # Row 33 — DISCRETIONARY header
        ["DISCRETIONARY  —  Where Leaks Happen", "", "", "", ""],
        # Row 34 — col headers
        ["Category", "Spent ($)", "Budget ($)", "vs Budget", "Pct of Income"],
        # Rows 35-46 — discretionary
        ["Dining Out / Restaurants", "", "", "=IF(C35=\"\",\"\",C35-B35)", "=IFERROR(B35/$B$10,0)"],
        ["Entertainment / Golf",     "", "", "=IF(C36=\"\",\"\",C36-B36)", "=IFERROR(B36/$B$10,0)"],
        ["Shopping",                 "", "", "=IF(C37=\"\",\"\",C37-B37)", "=IFERROR(B37/$B$10,0)"],
        ["Subscriptions",            "", "", "=IF(C38=\"\",\"\",C38-B38)", "=IFERROR(B38/$B$10,0)"],
        ["Streaming Services",       "", "", "=IF(C39=\"\",\"\",C39-B39)", "=IFERROR(B39/$B$10,0)"],
        ["Car Wash",                 "", "", "=IF(C40=\"\",\"\",C40-B40)", "=IFERROR(B40/$B$10,0)"],
        ["Gym / Fitness",            "", "", "=IF(C41=\"\",\"\",C41-B41)", "=IFERROR(B41/$B$10,0)"],
        ["Beauty / Personal",        "", "", "=IF(C42=\"\",\"\",C42-B42)", "=IFERROR(B42/$B$10,0)"],
        ["Travel",                   "", "", "=IF(C43=\"\",\"\",C43-B43)", "=IFERROR(B43/$B$10,0)"],
        ["Cash / ATM Withdrawals",   "", "", "=IF(C44=\"\",\"\",C44-B44)", "=IFERROR(B44/$B$10,0)"],
        ["Bank Fees (NSF)",          "", "", "=IF(C45=\"\",\"\",C45-B45)", "=IFERROR(B45/$B$10,0)"],
        ["Taxes / Tax Services",     "", "", "=IF(C46=\"\",\"\",C46-B46)", "=IFERROR(B46/$B$10,0)"],
        # Row 47 — subtotal discretionary
        ["SUBTOTAL — DISCRETIONARY", "=SUM(B35:B46)", "", "", "=IFERROR(B47/$B$10,0)"],
        # Row 48 — blank
        ["", "", "", "", ""],
        # Row 49 — SAVINGS header
        ["SAVINGS + EXTRA DEBT", "", "", "", ""],
        # Row 50 — col headers
        ["Category", "Spent ($)", "Budget ($)", "vs Budget", "Pct of Income"],
        # Rows 51-53
        ["Extra Debt Payments",  "", "", "=IF(C51=\"\",\"\",C51-B51)", "=IFERROR(B51/$B$10,0)"],
        ["Savings Transfers",    "", "", "=IF(C52=\"\",\"\",C52-B52)", "=IFERROR(B52/$B$10,0)"],
        ["Investments",          "", "", "=IF(C53=\"\",\"\",C53-B53)", "=IFERROR(B53/$B$10,0)"],
        # Row 54 — subtotal savings
        ["SUBTOTAL — SAVINGS", "=SUM(B51:B53)", "", "", "=IFERROR(B54/$B$10,0)"],
        # Row 55 — blank
        ["", "", "", "", ""],
        # Row 56 — SUMMARY header
        ["SUMMARY", "", "", "", ""],
        # Rows 57-62
        ["Total Income",         "=B10",              "", "", ""],
        ["Fixed Expenses",       "=B24",              "", "", "=IFERROR(B58/B57,0)"],
        ["Essentials",           "=B31",              "", "", "=IFERROR(B59/B57,0)"],
        ["Discretionary",        "=B47",              "", "", "=IFERROR(B60/B57,0)"],
        ["Savings + Extra Debt", "=B54",              "", "", "=IFERROR(B61/B57,0)"],
        ["TOTAL SPENT",          "=B58+B59+B60+B61",  "", "", "=IFERROR(B62/B57,0)"],
        # Row 63 — blank
        ["", "", "", "", ""],
        # Row 64 — MONEY LEFT OVER
        ["MONEY LEFT OVER",      "=B57-B62",          "", "", ""],
        # Row 65 — savings rate
        ["Savings Rate",         "=IFERROR(B54/B10,0)", "", "", ""],
    ]
    return rows


def sanitize_rows(rows):
    """Strip cmd-special chars from string cells (not formulas)."""
    safe = []
    for row in rows:
        safe_row = []
        for cell in row:
            if isinstance(cell, str) and not cell.startswith("="):
                cell = cell.replace("&", "+").replace("%", "pct")
            safe_row.append(cell)
        safe.append(safe_row)
    return safe

def gws_values(tab_name, rows):
    payload = {
        "valueInputOption": "USER_ENTERED",
        "data": [{"range": f"'{tab_name}'!A1", "values": rows}]
    }
    params = {"spreadsheetId": SPREADSHEET_ID}
    # Write in chunks of 15 rows
    all_rows = sanitize_rows(rows)
    start = 1
    for i in range(0, len(all_rows), 15):
        chunk = all_rows[i:i+15]
        p = {"valueInputOption": "USER_ENTERED",
             "data": [{"range": f"'{tab_name}'!A{start}", "values": chunk}]}
        r = subprocess.run(
            ["gws.cmd", "sheets", "spreadsheets", "values", "batchUpdate",
             "--params", json.dumps({"spreadsheetId": SPREADSHEET_ID}),
             "--json",   json.dumps(p, ensure_ascii=False)],
            capture_output=True, text=True, encoding="utf-8"
        )
        if r.returncode != 0:
            return False, r.stderr[:200]
        start += len(chunk)
    return True, ""


def gws_format(requests):
    """Send formatting requests in chunks of 8."""
    params = {"spreadsheetId": SPREADSHEET_ID}
    for i in range(0, len(requests), 8):
        chunk = requests[i:i+8]
        r = subprocess.run(
            ["gws.cmd", "sheets", "spreadsheets", "batchUpdate",
             "--params", json.dumps(params),
             "--json",   json.dumps({"requests": chunk}, ensure_ascii=False)],
            capture_output=True, text=True, encoding="utf-8"
        )
        if r.returncode != 0:
            return False, r.stderr[:200]
    return True, ""


def clear_tab(tab_name):
    subprocess.run(
        ["gws.cmd", "sheets", "spreadsheets", "values", "clear",
         "--params", json.dumps({
             "spreadsheetId": SPREADSHEET_ID,
             "range": f"'{tab_name}'!A1:Z200"
         })],
        capture_output=True, text=True, encoding="utf-8"
    )


def format_tab(sheet_id):
    navy    = {"red": 0.118, "green": 0.227, "blue": 0.373}
    red_hdr = {"red": 0.600, "green": 0.100, "blue": 0.100}  # dark red for discretionary
    grn_hdr = {"red": 0.063, "green": 0.373, "blue": 0.200}  # dark green for savings
    blue_md = {"red": 0.145, "green": 0.388, "blue": 0.922}
    blue_lt = {"red": 0.859, "green": 0.918, "blue": 0.996}
    white   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
    gray_lt = {"red": 0.953, "green": 0.957, "blue": 0.965}
    gray_md = {"red": 0.878, "green": 0.878, "blue": 0.878}
    yellow  = {"red": 0.996, "green": 0.988, "blue": 0.910}
    blue_in = {"red": 0.114, "green": 0.306, "blue": 0.847}
    alt1    = {"red": 0.976, "green": 0.980, "blue": 0.984}
    pos_bg  = {"red": 0.820, "green": 0.980, "blue": 0.898}
    neg_bg  = {"red": 0.996, "green": 0.886, "blue": 0.886}
    blk     = {"red": 0.0,   "green": 0.0,   "blue": 0.0}
    green_t = {"red": 0.063, "green": 0.373, "blue": 0.200}

    def rng(r1, c1, r2, c2):
        return {"sheetId": sheet_id,
                "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    def repeat(r1, c1, r2, c2, fmt, fields):
        return {"repeatCell": {"range": rng(r1,c1,r2,c2),
                               "cell": {"userEnteredFormat": fmt}, "fields": fields}}

    flds_all  = "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"
    flds_bg   = "userEnteredFormat.backgroundColor"
    flds_num  = "userEnteredFormat(numberFormat,horizontalAlignment)"
    flds_txt  = "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)"

    def hdr(r, color, text_size=11):
        return repeat(r, 0, r+1, 5, {
            "backgroundColor": color,
            "textFormat": {"bold": True, "fontSize": text_size, "foregroundColor": white},
            "horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE",
        }, flds_all)

    def col_hdr(r):
        return repeat(r, 0, r+1, 5, {
            "backgroundColor": blue_lt,
            "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": navy},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }, flds_all)

    def subtotal(r):
        return repeat(r, 0, r+1, 5, {
            "backgroundColor": gray_md,
            "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": blk},
            "verticalAlignment": "MIDDLE",
        }, flds_all)

    def data_rows(r1, r2):
        return [
            repeat(r, 0, r+1, 5, {
                "backgroundColor": alt1 if i % 2 == 0 else white,
                "textFormat": {"fontSize": 10},
                "verticalAlignment": "MIDDLE",
            }, flds_all)
            for i, r in enumerate(range(r1, r2))
        ]

    requests = [
        # Freeze top 2 rows + col A
        {"updateSheetProperties": {
            "properties": {"sheetId": sheet_id,
                           "gridProperties": {"frozenRowCount": 2, "frozenColumnCount": 1}},
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
        }},
        # Row 1 — title (navy, large)
        hdr(0, navy, 14),
        # Row 2 — snapshot (medium blue)
        hdr(1, blue_md, 10),
        # Section headers
        hdr(3, navy, 11),          # INCOME
        hdr(11, navy, 11),         # FIXED EXPENSES
        hdr(25, navy, 11),         # ESSENTIALS
        hdr(32, red_hdr, 11),      # DISCRETIONARY
        hdr(48, grn_hdr, 11),      # SAVINGS
        hdr(55, navy, 11),         # SUMMARY
        # Column header rows
        col_hdr(4),   # income cols
        col_hdr(12),  # fixed cols
        col_hdr(26),  # essentials cols
        col_hdr(33),  # discretionary cols
        col_hdr(49),  # savings cols
        # Subtotal rows
        subtotal(9),   # TOTAL INCOME
        subtotal(23),  # SUBTOTAL FIXED
        subtotal(30),  # SUBTOTAL ESSENTIALS
        subtotal(46),  # SUBTOTAL DISCRETIONARY
        subtotal(53),  # SUBTOTAL SAVINGS
        subtotal(61),  # TOTAL SPENT
        # Summary section rows
        *[repeat(r, 0, r+1, 5, {
            "backgroundColor": gray_lt,
            "textFormat": {"fontSize": 10, "bold": False},
            "verticalAlignment": "MIDDLE",
        }, flds_all) for r in range(56, 62)],
        # MONEY LEFT OVER — row 64 (index 63) — big, navy bg
        repeat(63, 0, 64, 5, {
            "backgroundColor": navy,
            "textFormat": {"bold": True, "fontSize": 16, "foregroundColor": white},
            "horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE",
        }, flds_all),
        # Savings rate row
        repeat(64, 0, 65, 5, {
            "backgroundColor": gray_lt,
            "textFormat": {"fontSize": 10},
            "verticalAlignment": "MIDDLE",
        }, flds_all),
        # Alternating data rows
        *data_rows(5, 9),    # income data
        *data_rows(13, 23),  # fixed data
        *data_rows(27, 30),  # essentials data
        *data_rows(34, 46),  # discretionary data
        *data_rows(50, 53),  # savings data
        # Column widths: A=230, B=110, C=110, D=100, E=95
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 230}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 1, "endIndex": 3},
            "properties": {"pixelSize": 115}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 3, "endIndex": 4},
            "properties": {"pixelSize": 105}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": 4, "endIndex": 5},
            "properties": {"pixelSize": 95}, "fields": "pixelSize"
        }},
        # Row heights
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": 0, "endIndex": 2},
            "properties": {"pixelSize": 30}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": 2, "endIndex": 65},
            "properties": {"pixelSize": 22}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": 63, "endIndex": 64},
            "properties": {"pixelSize": 40}, "fields": "pixelSize"
        }},
        # Currency format on B and C columns
        {"repeatCell": {
            "range": rng(5, 1, 65, 3),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00;($#,##0.00);\"-\""},
                "horizontalAlignment": "RIGHT",
            }},
            "fields": flds_num
        }},
        # +/- column D: currency with color via conditional formatting
        {"repeatCell": {
            "range": rng(5, 3, 65, 4),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00;($#,##0.00);\"-\""},
                "horizontalAlignment": "RIGHT",
            }},
            "fields": flds_num
        }},
        # % Income column E
        {"repeatCell": {
            "range": rng(5, 4, 65, 5),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "0.0%"},
                "horizontalAlignment": "CENTER",
            }},
            "fields": flds_num
        }},
        # Budget column C — yellow input cells for data rows
        {"repeatCell": {
            "range": rng(5, 2, 9, 3),    # income budget
            "cell": {"userEnteredFormat": {
                "backgroundColor": yellow,
                "textFormat": {"foregroundColor": blue_in},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }},
        {"repeatCell": {
            "range": rng(13, 2, 23, 3),   # fixed budget
            "cell": {"userEnteredFormat": {
                "backgroundColor": yellow,
                "textFormat": {"foregroundColor": blue_in},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }},
        {"repeatCell": {
            "range": rng(27, 2, 30, 3),   # essentials budget
            "cell": {"userEnteredFormat": {
                "backgroundColor": yellow,
                "textFormat": {"foregroundColor": blue_in},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }},
        {"repeatCell": {
            "range": rng(34, 2, 46, 3),   # discretionary budget
            "cell": {"userEnteredFormat": {
                "backgroundColor": yellow,
                "textFormat": {"foregroundColor": blue_in},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }},
        {"repeatCell": {
            "range": rng(50, 2, 53, 3),   # savings budget
            "cell": {"userEnteredFormat": {
                "backgroundColor": yellow,
                "textFormat": {"foregroundColor": blue_in},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }},
        # Conditional format: +/- D column — green if positive (under budget), red if over
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [rng(5, 3, 65, 4)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_GREATER",
                                  "values": [{"userEnteredValue": "0"}]},
                    "format": {"backgroundColor": pos_bg,
                               "textFormat": {"foregroundColor": green_t}}
                }
            }, "index": 0
        }},
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [rng(5, 3, 65, 4)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_LESS",
                                  "values": [{"userEnteredValue": "0"}]},
                    "format": {"backgroundColor": neg_bg,
                               "textFormat": {"foregroundColor": {"red":0.6,"green":0,"blue":0}}}
                }
            }, "index": 1
        }},
        # % Income column — red if > 20% of income (potential leak)
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [rng(34, 4, 46, 5)],  # discretionary % only
                "booleanRule": {
                    "condition": {"type": "NUMBER_GREATER",
                                  "values": [{"userEnteredValue": "0.1"}]},
                    "format": {"backgroundColor": neg_bg}
                }
            }, "index": 2
        }},
        # Borders
        {"updateBorders": {
            "range": rng(0, 0, 65, 5),
            "innerHorizontal": {"style": "SOLID",
                                "color": {"red": 0.85, "green": 0.85, "blue": 0.85}},
            "innerVertical":   {"style": "SOLID",
                                "color": {"red": 0.85, "green": 0.85, "blue": 0.85}},
        }},
    ]
    return gws_format(requests)


def main():
    print("Setting up monthly tabs...")
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--start-month", type=int, default=1)
    args = parser.parse_args()

    for (month_num, tab_name, month_name, sheet_id) in MONTHS:
        if month_num < args.start_month:
            continue
        print(f"  {tab_name}...", end=" ", flush=True)

        # Clear existing content
        clear_tab(tab_name)

        # Write data
        rows = tab_values(month_name, 2026)
        ok, err = gws_values(tab_name, rows)
        if not ok:
            print(f"ERROR writing: {err}")
            continue

        # Apply formatting (with small delay to avoid quota)
        time.sleep(2)
        ok, err = format_tab(sheet_id)
        if not ok:
            print(f"ERROR formatting: {err}")
        else:
            print("OK")

        time.sleep(8)

    print("\nDone. Re-run the import script to repopulate actuals:")
    print("  python scripts/import_bank_csv.py \"C:/Users/Bruce/Downloads/303482_S9.csv\"")
    print(f"\nSheet: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()
