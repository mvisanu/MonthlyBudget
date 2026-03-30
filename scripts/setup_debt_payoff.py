#!/usr/bin/env python3
"""
Builds a full debt payoff tracker with avalanche + snowball strategies.

Usage:
    python scripts/setup_debt_payoff.py
"""

import io, json, subprocess, sys, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SPREADSHEET_ID = "1DEaFJvnXOM_B9GglT6sKXZzq0lCYF__EP_reMm3qu-w"
DEBT_SHEET_ID  = 698513475
GWS_SCRIPT     = r"C:\Users\Bruce\AppData\Roaming\npm\node_modules\@googleworkspace\cli\run-gws.js"

# ── Debts: [Name, Type, Balance, APR%, MinPmt] ───────────────────────────────
# Sorted lowest-balance-first (Snowball order by default)
# User should fill in APR% and Min Payment — estimates provided where known
DEBTS = [
    # Name                  Type            Balance   APR%   MinPmt
    ["First Access",        "Credit Card",  8.99,     29.99, 25.00],
    ["Indigo",              "Credit Card",  10.13,    29.99, 25.00],
    ["Amex 7697",           "Credit Card",  739.83,   27.99, 25.00],
    ["Sams Club",           "Credit Card",  817.75,   29.99, 25.00],
    ["Capital One Savor",   "Credit Card",  888.39,   29.99, 25.00],
    ["X5 Visa 7213",        "Credit Card",  1007.37,  27.99, 30.00],
    ["Venture",             "Credit Card",  1165.14,  29.99, 30.00],
    ["Affirm",              "BNPL / Loan",  3996.73,  0.00,  150.00],
    ["Amex 9939",           "Credit Card",  3197.65,  27.99, 65.00],
    ["Bass Pro",            "Credit Card",  4994.37,  29.99, 100.00],
]

# Row indices (1-based) for each debt in the sheet — used in formulas
DEBT_START_ROW = 6   # first debt is row 6
DEBT_END_ROW   = DEBT_START_ROW + len(DEBTS) - 1  # row 15


def gws_call(cmd_args, payload):
    r = subprocess.run(
        ["node", GWS_SCRIPT] + cmd_args +
        ["--params", json.dumps({"spreadsheetId": SPREADSHEET_ID}),
         "--json",   json.dumps(payload, ensure_ascii=False)],
        capture_output=True, text=True, encoding="utf-8"
    )
    return r.returncode == 0, r.stderr[:300] if r.returncode != 0 else ""


def rng(r1, c1, r2, c2):
    return {"sheetId": DEBT_SHEET_ID, "startRowIndex": r1, "endRowIndex": r2,
            "startColumnIndex": c1, "endColumnIndex": c2}


def rep(r1, c1, r2, c2, fmt, fields=None):
    return {"repeatCell": {
        "range": rng(r1, c1, r2, c2),
        "cell": {"userEnteredFormat": fmt},
        "fields": fields or "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)"
    }}


def build_rows():
    """Build the full sheet data."""
    #
    # Layout:
    #  Row  1 : Title
    #  Row  2 : Subtitle / instructions
    #  Row  3 : blank
    #  Row  4 : DEBT INVENTORY header
    #  Row  5 : col headers
    #  Rows 6-15 : each debt (10 debts)
    #  Row 16 : blank
    #  Row 17 : TOTALS
    #  Row 18 : Total Debt
    #  Row 19 : Total Min Payments
    #  Row 20 : Extra Payment Available (from budget)
    #  Row 21 : blank
    #  Row 22 : STRATEGY COMPARISON header
    #  Row 23 : col headers
    #  Row 24 : Snowball (lowest balance first) — estimated payoff
    #  Row 25 : Avalanche (highest APR first) — estimated payoff
    #  Row 26 : blank
    #  Row 27 : SNOWBALL ORDER header
    #  Row 28 : col headers
    #  Rows 29-38 : snowball order (same as DEBTS — already sorted low-to-high)
    #  Row 39 : blank
    #  Row 40 : AVALANCHE ORDER header
    #  Row 41 : col headers
    #  Rows 42-51 : avalanche order (sort by APR% desc — user fills in real APRs)
    #  Row 52 : blank
    #  Row 53 : NOTE
    #
    # Cols: A=Creditor, B=Type, C=Balance, D=APR%, E=Min Pmt, F=Months to Payoff, G=Est Payoff Date, H=Notes

    nper = lambda r: (
        f'=IF(AND(C{r}>0,E{r}>0,D{r}>0),'
        f'ROUNDUP(NPER(D{r}/100/12,-E{r},C{r}),0),'
        f'IF(AND(C{r}>0,E{r}>0),"0% - check APR","Fill in"))'
    )
    payoff_date = lambda r: (
        f'=IF(ISNUMBER(F{r}),TEXT(EDATE(TODAY(),F{r}),"mmm yyyy"),"")'
    )

    col_headers = ["Creditor", "Type", "Balance ($)", "APR %", "Min Payment ($)",
                   "Months to Pay Off", "Est. Payoff Date", "Notes"]

    rows = []

    # Row 1
    rows.append(["DEBT PAYOFF TRACKER", "", "", "", "", "", "", ""])
    # Row 2
    rows.append(["Fill in real APR% and Min Payment to activate payoff formulas. Yellow = needs your input.", "", "", "", "", "", "", ""])
    # Row 3
    rows.append(["", "", "", "", "", "", "", ""])
    # Row 4
    rows.append(["DEBT INVENTORY (as of March 2026)", "", "", "", "", "", "", ""])
    # Row 5
    rows.append(col_headers)

    # Rows 6-15 — debts
    for i, (name, dtype, bal, apr, minpmt) in enumerate(DEBTS):
        r = DEBT_START_ROW + i
        rows.append([name, dtype, bal, apr, minpmt, nper(r), payoff_date(r), ""])

    # Row 16 — blank
    rows.append(["", "", "", "", "", "", "", ""])
    # Row 17 — TOTALS header
    rows.append(["TOTALS", "", "", "", "", "", "", ""])
    # Row 18 — Total debt
    rows.append(["Total Outstanding Debt", "",
                 f"=SUM(C{DEBT_START_ROW}:C{DEBT_END_ROW})", "", "", "", "", ""])
    # Row 19 — Total minimums
    rows.append(["Total Minimum Payments / mo", "",
                 f"=SUM(E{DEBT_START_ROW}:E{DEBT_END_ROW})", "", "", "", "", ""])
    # Row 20 — Extra payment
    rows.append(["Extra Payment Available / mo", "", 500, "", "",
                 "(adjust this — from your monthly budget surplus)", "", ""])
    # Row 21 — Total monthly toward debt
    rows.append(["Total Toward Debt / mo", "",
                 "=C19+C20", "", "", "", "", ""])
    # Row 22 — blank
    rows.append(["", "", "", "", "", "", "", ""])

    # Row 23 — STRATEGY header
    rows.append(["PAYOFF STRATEGY COMPARISON", "", "", "", "", "", "", ""])
    # Row 24 — col headers
    rows.append(["Method", "Order", "How It Works", "Psychological Benefit", "", "", "", ""])
    # Row 25 — Snowball
    rows.append(["Snowball", "Lowest balance first",
                 "Pay minimums on all, throw extra at smallest debt",
                 "Quick wins — see debts disappear fast", "", "", "", ""])
    # Row 26 — Avalanche
    rows.append(["Avalanche", "Highest APR% first",
                 "Pay minimums on all, throw extra at highest-rate debt",
                 "Saves the most money in interest over time", "", "", "", ""])
    # Row 27 — blank
    rows.append(["", "", "", "", "", "", "", ""])

    # Row 28 — SNOWBALL ORDER header
    rows.append(["SNOWBALL ORDER — Attack Lowest Balance First", "", "", "", "", "", "", ""])
    # Row 29 — col headers
    rows.append(col_headers)

    # Rows 30-39 — snowball order (already sorted low→high)
    snowball = sorted(DEBTS, key=lambda x: x[2])
    for i, (name, dtype, bal, apr, minpmt) in enumerate(snowball):
        priority = i + 1
        rows.append([f"#{priority}  {name}", dtype, bal, apr, minpmt,
                     f'=IF(AND({bal}>0,{minpmt}>0,{apr}>0),ROUNDUP(NPER({apr}/100/12,-{minpmt},{bal}),0),"Fill in")',
                     f'=IF(ISNUMBER(F{29+i}),TEXT(EDATE(TODAY(),F{29+i}),"mmm yyyy"),"")',
                     "TARGET" if priority == 1 else ""])

    # blank
    rows.append(["", "", "", "", "", "", "", ""])

    # AVALANCHE ORDER header
    rows.append(["AVALANCHE ORDER — Attack Highest APR% First", "", "", "", "", "", "", ""])
    # col headers
    rows.append(col_headers)

    # Avalanche order (high→low APR)
    avalanche = sorted(DEBTS, key=lambda x: x[3], reverse=True)
    for i, (name, dtype, bal, apr, minpmt) in enumerate(avalanche):
        priority = i + 1
        rows.append([f"#{priority}  {name}", dtype, bal, apr, minpmt,
                     f'=IF(AND({bal}>0,{minpmt}>0,{apr}>0),ROUNDUP(NPER({apr}/100/12,-{minpmt},{bal}),0),"Fill in")',
                     f'=IF(ISNUMBER(F{42+i}),TEXT(EDATE(TODAY(),F{42+i}),"mmm yyyy"),"")',
                     "TARGET" if priority == 1 else ""])

    # blank
    rows.append(["", "", "", "", "", "", "", ""])

    # NOTE
    rows.append(["NOTE: Payoff dates assume no new charges. Update balances monthly as you pay down.", "", "", "", "", "", "", ""])
    rows.append(["Affirm APR shown as 0% — confirm your actual Affirm loan rate (may vary 0-36%).", "", "", "", "", "", "", ""])

    return rows


def write_values(rows):
    # Clear first
    gws_call(["sheets", "spreadsheets", "values", "clear"],
             {"range": chr(0x1F4B3) + " Debt Payoff!A1:H60"})
    time.sleep(0.5)

    start = 1
    for i in range(0, len(rows), 5):
        chunk = rows[i:i+5]
        payload = {
            "valueInputOption": "USER_ENTERED",
            "data": [{"range": chr(0x1F4B3) + f" Debt Payoff!A{start}", "values": chunk}]
        }
        ok, err = gws_call(["sheets", "spreadsheets", "values", "batchUpdate"], payload)
        if not ok:
            print(f"  Write error at row {start}: {err}")
            return False
        start += len(chunk)
        time.sleep(0.3)
    return True


def apply_formatting():
    navy    = {"red": 0.118, "green": 0.227, "blue": 0.373}
    white   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
    red_hdr = {"red": 0.698, "green": 0.133, "blue": 0.133}
    grn_hdr = {"red": 0.180, "green": 0.490, "blue": 0.196}
    blue_lt = {"red": 0.859, "green": 0.918, "blue": 0.996}
    yel     = {"red": 1.000, "green": 0.973, "blue": 0.820}
    grn_lt  = {"red": 0.851, "green": 0.953, "blue": 0.867}
    red_lt  = {"red": 0.996, "green": 0.886, "blue": 0.886}
    gray_md = {"red": 0.878, "green": 0.878, "blue": 0.878}
    gray_lt = {"red": 0.953, "green": 0.957, "blue": 0.965}
    alt     = {"red": 0.976, "green": 0.980, "blue": 0.984}
    blk     = {"red": 0.0,   "green": 0.0,   "blue": 0.0}
    amber   = {"red": 0.737, "green": 0.604, "blue": 0.118}

    def hdr(r, bg, size=11):
        return rep(r, 0, r+1, 8, {
            "backgroundColor": bg,
            "textFormat": {"bold": True, "fontSize": size, "foregroundColor": white},
            "horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE",
        })

    def col_hdr(r):
        return rep(r, 0, r+1, 8, {
            "backgroundColor": {"red": 0.267, "green": 0.431, "blue": 0.643},
            "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": white},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        })

    def data_rows(r1, r2, bg1, bg2):
        reqs = []
        for i, r in enumerate(range(r1, r2)):
            reqs.append(rep(r, 0, r+1, 8, {
                "backgroundColor": bg1 if i % 2 == 0 else bg2,
                "textFormat": {"fontSize": 10},
                "verticalAlignment": "MIDDLE",
            }))
        return reqs

    def subtotal(r):
        return rep(r, 0, r+1, 8, {
            "backgroundColor": gray_md,
            "textFormat": {"bold": True, "fontSize": 10},
            "verticalAlignment": "MIDDLE",
        })

    requests = [
        # Column widths
        {"updateDimensionProperties": {"range": {"sheetId": DEBT_SHEET_ID, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 200}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": DEBT_SHEET_ID, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 110}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": DEBT_SHEET_ID, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 120}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": DEBT_SHEET_ID, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 80}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": DEBT_SHEET_ID, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5}, "properties": {"pixelSize": 130}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": DEBT_SHEET_ID, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 6}, "properties": {"pixelSize": 130}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": DEBT_SHEET_ID, "dimension": "COLUMNS", "startIndex": 6, "endIndex": 7}, "properties": {"pixelSize": 120}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": DEBT_SHEET_ID, "dimension": "COLUMNS", "startIndex": 7, "endIndex": 8}, "properties": {"pixelSize": 220}, "fields": "pixelSize"}},
        # Row 1 — title (navy, big)
        hdr(0, navy, 14),
        # Row 2 — instructions (amber)
        rep(1, 0, 2, 8, {
            "backgroundColor": yel,
            "textFormat": {"fontSize": 9, "italic": True, "foregroundColor": {"red":0.4,"green":0.3,"blue":0.0}},
            "verticalAlignment": "MIDDLE",
        }),
        # Row 4 — DEBT INVENTORY header (dark red)
        hdr(3, red_hdr, 11),
        # Row 5 — col headers
        col_hdr(4),
        # Rows 6-15 — debt data (alternating red tint)
        *data_rows(5, 15, red_lt, white),
        # Row 17 — TOTALS header
        hdr(16, red_hdr, 11),
        # Rows 18-21 — totals
        *[subtotal(r) for r in [17, 18, 19, 20]],
        # Row 23 — STRATEGY header (navy)
        hdr(22, navy, 11),
        # Rows 24-26 — strategy rows
        col_hdr(23),
        *data_rows(24, 27, grn_lt, white),
        # Row 28 — SNOWBALL header (green)
        hdr(27, grn_hdr, 11),
        # Row 29 — col headers
        col_hdr(28),
        # Rows 30-39 — snowball debts
        *data_rows(29, 39, grn_lt, white),
        # Row 41 — AVALANCHE header (amber/dark)
        hdr(40, amber, 11),
        # Row 42 — col headers
        col_hdr(41),
        # Rows 43-52 — avalanche debts
        *data_rows(42, 52, {"red":1.0,"green":0.973,"blue":0.820}, white),
        # Currency format: C col (balance) and E col (min pmt) across all debt rows
        rep(5, 2, 52, 3, {
            "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00"},
            "horizontalAlignment": "RIGHT",
        }, fields="userEnteredFormat(numberFormat,horizontalAlignment)"),
        rep(5, 4, 52, 5, {
            "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00"},
            "horizontalAlignment": "RIGHT",
        }, fields="userEnteredFormat(numberFormat,horizontalAlignment)"),
        # APR% col D — percent format
        rep(5, 3, 52, 4, {
            "numberFormat": {"type": "NUMBER", "pattern": "0.00%"},
            "horizontalAlignment": "CENTER",
        }, fields="userEnteredFormat(numberFormat,horizontalAlignment)"),
        # Freeze rows 1-2
        {"updateSheetProperties": {
            "properties": {"sheetId": DEBT_SHEET_ID, "gridProperties": {"frozenRowCount": 2}},
            "fields": "gridProperties.frozenRowCount"
        }},
        # Highlight "TARGET" cell in Notes col for snowball row 1
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [rng(29, 7, 39, 8)],
                "booleanRule": {
                    "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "TARGET"}]},
                    "format": {
                        "backgroundColor": {"red":1.0,"green":0.843,"blue":0.0},
                        "textFormat": {"bold": True, "foregroundColor": blk}
                    }
                }
            }, "index": 0
        }},
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [rng(42, 7, 52, 8)],
                "booleanRule": {
                    "condition": {"type": "TEXT_EQ", "values": [{"userEnteredValue": "TARGET"}]},
                    "format": {
                        "backgroundColor": {"red":1.0,"green":0.843,"blue":0.0},
                        "textFormat": {"bold": True, "foregroundColor": blk}
                    }
                }
            }, "index": 1
        }},
    ]

    for i in range(0, len(requests), 8):
        chunk = requests[i:i+8]
        ok, err = gws_call(["sheets", "spreadsheets", "batchUpdate"], {"requests": chunk})
        if not ok:
            print(f"  Format error (chunk {i}): {err}")
            return False
        time.sleep(0.8)
    return True


def main():
    print("Building Debt Payoff tab...")
    rows = build_rows()
    print(f"  Writing {len(rows)} rows...", end=" ", flush=True)
    if not write_values(rows):
        return
    print("OK")
    print("  Applying formatting...", end=" ", flush=True)
    if apply_formatting():
        print("OK")
    else:
        print("ERRORS (check above)")

    print(f"\nDone.")
    print(f"  APR% column (D) has estimated rates — update with your actual rates.")
    print(f"  Min Payment (E) has estimates — update with your actual minimums.")
    print(f"  Extra payment in row 20 defaults to $500 — adjust to your real surplus.")
    print(f"  https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()
