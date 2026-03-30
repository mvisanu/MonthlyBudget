#!/usr/bin/env python3
"""
Rewrites the Dashboard tab with auto-updating income summary + pie chart.

Usage:
    python scripts/setup_dashboard.py
"""

import io, json, subprocess, sys, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SPREADSHEET_ID = "1DEaFJvnXOM_B9GglT6sKXZzq0lCYF__EP_reMm3qu-w"
DASHBOARD_SHEET_ID = 611629967

# CHOOSE formula helper — converts $B$2 (1-12) to tab name
# Avoids & concatenation by embedding full ref strings
CHOOSE_TAB = (
    "CHOOSE($B$2,"
    "\"'\\U0001f4c5 Jan'\","
    "\"'\\U0001f4c5 Feb'\","
    "\"'\\U0001f4c5 Mar'\","
    "\"'\\U0001f4c5 Apr'\","
    "\"'\\U0001f4c5 May'\","
    "\"'\\U0001f4c5 Jun'\","
    "\"'\\U0001f4c5 Jul'\","
    "\"'\\U0001f4c5 Aug'\","
    "\"'\\U0001f4c5 Sep'\","
    "\"'\\U0001f4c5 Oct'\","
    "\"'\\U0001f4c5 Nov'\","
    "\"'\\U0001f4c5 Dec'\")"
)

def indirect(cell):
    """=IFERROR(INDIRECT(CHOOSE(...)&"!BXX"),0)"""
    return f'=IFERROR(INDIRECT(CHOOSE($B$2,"\'📅 Jan\'","\'📅 Feb\'","\'📅 Mar\'","\'📅 Apr\'","\'📅 May\'","\'📅 Jun\'","\'📅 Jul\'","\'📅 Aug\'","\'📅 Sep\'","\'📅 Oct\'","\'📅 Nov\'","\'📅 Dec\'")&"!{cell}"),0)'

def month_name_formula():
    return '=CHOOSE($B$2,"January","February","March","April","May","June","July","August","September","October","November","December")'

# ── Dashboard layout ────────────────────────────────────────────────────────
#
# Cols A-C: Summary panel (left side)
# Cols E-G: Pie chart data (feeds the chart, can be hidden)
# Cols H+:  Chart sits here (embedded)
#
# Row  1: Title
# Row  2: "Active Month:" | [1-12 input] | [month name formula]
# Row  3: blank
# Row  4: INCOME (section header)
# Row  5: col headers: Source | Amount
# Row  6: Paycheck 1 (Humana)
# Row  7: Paycheck 2 (Boeing)
# Row  8: Freelance / Side Hustle
# Row  9: Other Income
# Row 10: TOTAL INCOME
# Row 11: blank
# Row 12: SPENDING SUMMARY (section header)
# Row 13: col headers
# Row 14: Fixed Expenses
# Row 15: Groceries
# Row 16: Fuel / Gas
# Row 17: Medical
# Row 18: Dining Out
# Row 19: Entertainment / Golf   ← leak indicator
# Row 20: Shopping
# Row 21: Subscriptions
# Row 22: Streaming
# Row 23: Car Wash
# Row 24: Gym / Fitness
# Row 25: Beauty / Personal
# Row 26: Travel
# Row 27: Cash / ATM
# Row 28: Bank Fees (NSF)
# Row 29: Taxes
# Row 30: Savings + Extra Debt
# Row 31: TOTAL SPENT
# Row 32: blank
# Row 33: MONEY LEFT OVER  (BIG)
# Row 34: Savings Rate
# Row 35: blank
# Rows 37-53: Pie chart data (cols E-F) — category/amount pairs

ROWS = [
    # Row 1 — Title
    ["Personal Finance Dashboard", "", "", "", "Category", "Amount ($)"],
    # Row 2 — Month selector
    ["Active Month (1-12):", 2, month_name_formula(), "", "", ""],
    # Row 3 — blank
    ["", "", "", "", "", ""],
    # Row 4 — INCOME header
    ["INCOME", "", "", "", "", ""],
    # Row 5 — col headers
    ["Source", "Amount ($)", "Pct of Income", "", "", ""],
    # Row 6-9 — income sources
    ["Paycheck 1 (Humana)",    indirect("B6"),  f"=IFERROR(B6/B10,0)", "", "", ""],
    ["Paycheck 2 (Boeing)",    indirect("B7"),  f"=IFERROR(B7/B10,0)", "", "", ""],
    ["Freelance / Side Hustle",indirect("B8"),  f"=IFERROR(B8/B10,0)", "", "", ""],
    ["Other Income",           indirect("B9"),  f"=IFERROR(B9/B10,0)", "", "", ""],
    # Row 10 — total income
    ["TOTAL INCOME", "=SUM(B6:B9)", "", "", "", ""],
    # Row 11 — blank
    ["", "", "", "", "", ""],
    # Row 12 — SPENDING header
    ["SPENDING BREAKDOWN", "", "", "", "", ""],
    # Row 13 — col headers
    ["Category", "Amount ($)", "Pct of Income", "", "", ""],
    # Rows 14-30 — spending categories (also write to E-F for pie chart)
    ["Fixed Expenses",       indirect("B24"), "=IFERROR(B14/B10,0)", "", "Fixed Expenses",       indirect("B24")],
    ["Groceries",            indirect("B28"), "=IFERROR(B15/B10,0)", "", "Groceries",            indirect("B28")],
    ["Fuel / Gas",           indirect("B29"), "=IFERROR(B16/B10,0)", "", "Fuel / Gas",           indirect("B29")],
    ["Medical",              indirect("B30"), "=IFERROR(B17/B10,0)", "", "Medical",              indirect("B30")],
    ["Dining Out",           indirect("B35"), "=IFERROR(B18/B10,0)", "", "Dining Out",           indirect("B35")],
    ["Entertainment / Golf", indirect("B36"), "=IFERROR(B19/B10,0)", "", "Entertainment / Golf", indirect("B36")],
    ["Shopping",             indirect("B37"), "=IFERROR(B20/B10,0)", "", "Shopping",             indirect("B37")],
    ["Subscriptions",        indirect("B38"), "=IFERROR(B21/B10,0)", "", "Subscriptions",        indirect("B38")],
    ["Streaming",            indirect("B39"), "=IFERROR(B22/B10,0)", "", "Streaming",            indirect("B39")],
    ["Car Wash",             indirect("B40"), "=IFERROR(B23/B10,0)", "", "Car Wash",             indirect("B40")],
    ["Gym / Fitness",        indirect("B41"), "=IFERROR(B24_/B10,0)","", "Gym / Fitness",        indirect("B41")],
    ["Beauty / Personal",    indirect("B42"), "=IFERROR(B25/B10,0)", "", "Beauty / Personal",    indirect("B42")],
    ["Travel",               indirect("B43"), "=IFERROR(B26/B10,0)", "", "Travel",               indirect("B43")],
    ["Cash / ATM",           indirect("B44"), "=IFERROR(B27/B10,0)", "", "Cash / ATM",           indirect("B44")],
    ["Bank Fees (NSF)",      indirect("B45"), "=IFERROR(B28/B10,0)", "", "Bank Fees (NSF)",      indirect("B45")],
    ["Taxes",                indirect("B46"), "=IFERROR(B29/B10,0)", "", "Taxes",                indirect("B46")],
    ["Savings + Extra Debt", indirect("B54"), "=IFERROR(B30/B10,0)", "", "Savings + Extra Debt", indirect("B54")],
    # Row 31 — Total spent
    ["TOTAL SPENT", "=SUM(B14:B30)", "=IFERROR(B31/B10,0)", "", "", ""],
    # Row 32 — blank
    ["", "", "", "", "", ""],
    # Row 33 — MONEY LEFT OVER
    ["MONEY LEFT OVER", "=B10-B31", "", "", "", ""],
    # Row 34 — savings rate
    ["Savings Rate", f"={indirect('B65')[1:]}", "", "", "", ""],
]

# Fix row 24 typo (B24_ should be B24)
for i, row in enumerate(ROWS):
    ROWS[i] = [c.replace("B24_", "B24") if isinstance(c, str) else c for c in row]


GWS_SCRIPT = r"C:\Users\Bruce\AppData\Roaming\npm\node_modules\@googleworkspace\cli\run-gws.js"

def gws_call(cmd_args, payload):
    # Call node directly (not gws.cmd) to bypass cmd.exe encoding issues with emoji.
    r = subprocess.run(
        ["node", GWS_SCRIPT] + cmd_args +
        ["--params", json.dumps({"spreadsheetId": SPREADSHEET_ID}),
         "--json",   json.dumps(payload, ensure_ascii=False)],
        capture_output=True, text=True, encoding="utf-8"
    )
    return r.returncode == 0, r.stderr[:300] if r.returncode != 0 else ""


def write_values(rows):
    """Write rows in chunks of 3 (formulas are long due to INDIRECT)."""
    start = 1
    for i in range(0, len(rows), 3):
        chunk = rows[i:i+3]
        payload = {
            "valueInputOption": "USER_ENTERED",
            "data": [{"range": f"'🏠 Dashboard'!A{start}", "values": chunk}]
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
    blue_md = {"red": 0.145, "green": 0.388, "blue": 0.922}
    blue_lt = {"red": 0.859, "green": 0.918, "blue": 0.996}
    white   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
    yellow  = {"red": 0.996, "green": 0.988, "blue": 0.910}
    blue_in = {"red": 0.114, "green": 0.306, "blue": 0.847}
    gray_lt = {"red": 0.953, "green": 0.957, "blue": 0.965}
    gray_md = {"red": 0.878, "green": 0.878, "blue": 0.878}
    alt1    = {"red": 0.976, "green": 0.980, "blue": 0.984}
    pos_bg  = {"red": 0.820, "green": 0.980, "blue": 0.898}
    neg_bg  = {"red": 0.996, "green": 0.886, "blue": 0.886}
    grn_t   = {"red": 0.063, "green": 0.373, "blue": 0.200}
    blk     = {"red": 0.0,   "green": 0.0,   "blue": 0.0}

    sid = DASHBOARD_SHEET_ID

    def rng(r1, c1, r2, c2):
        return {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2,
                "startColumnIndex": c1, "endColumnIndex": c2}

    def rep(r1, c1, r2, c2, fmt):
        return {"repeatCell": {
            "range": rng(r1,c1,r2,c2),
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)"
        }}

    def hdr(r, color, size=11):
        return rep(r, 0, r+1, 4, {
            "backgroundColor": color,
            "textFormat": {"bold": True, "fontSize": size, "foregroundColor": white},
            "horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE",
        })

    def subtotal_row(r):
        return rep(r, 0, r+1, 4, {
            "backgroundColor": gray_md,
            "textFormat": {"bold": True, "fontSize": 10},
            "verticalAlignment": "MIDDLE",
        })

    requests = [
        # Freeze rows 1-3
        {"updateSheetProperties": {
            "properties": {"sheetId": sid,
                           "gridProperties": {"frozenRowCount": 3}},
            "fields": "gridProperties.frozenRowCount"
        }},
        # Row 1 — title (navy, big)
        hdr(0, navy, 16),
        # Row 2 — month selector (blue)
        rep(1, 0, 2, 4, {
            "backgroundColor": blue_md,
            "textFormat": {"bold": True, "fontSize": 11, "foregroundColor": white},
            "verticalAlignment": "MIDDLE",
        }),
        # B2 — yellow input cell
        rep(1, 1, 2, 2, {
            "backgroundColor": yellow,
            "textFormat": {"bold": True, "fontSize": 14, "foregroundColor": blue_in},
            "horizontalAlignment": "CENTER",
        }),
        # C2 — month name (white text on blue)
        rep(1, 2, 2, 3, {
            "backgroundColor": blue_md,
            "textFormat": {"bold": True, "fontSize": 12, "foregroundColor": white},
            "horizontalAlignment": "LEFT",
        }),
        # Row 4 — INCOME header
        hdr(3, navy, 11),
        # Row 5 — col headers (light blue)
        rep(4, 0, 5, 4, {
            "backgroundColor": blue_lt,
            "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": navy},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }),
        # Rows 6-9 — income data (alternating)
        *[rep(r, 0, r+1, 4, {
            "backgroundColor": alt1 if i%2==0 else white,
            "textFormat": {"fontSize": 10},
            "verticalAlignment": "MIDDLE",
        }) for i, r in enumerate(range(5, 9))],
        # Row 10 — total income
        subtotal_row(9),
        # Row 12 — SPENDING header
        hdr(11, navy, 11),
        # Row 13 — col headers
        rep(12, 0, 13, 4, {
            "backgroundColor": blue_lt,
            "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": navy},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }),
        # Rows 14-30 — spending data (alternating)
        *[rep(r, 0, r+1, 4, {
            "backgroundColor": alt1 if i%2==0 else white,
            "textFormat": {"fontSize": 10},
            "verticalAlignment": "MIDDLE",
        }) for i, r in enumerate(range(13, 30))],
        # Row 31 — total spent
        subtotal_row(30),
        # Row 33 — MONEY LEFT OVER (big navy)
        rep(32, 0, 33, 4, {
            "backgroundColor": navy,
            "textFormat": {"bold": True, "fontSize": 18, "foregroundColor": white},
            "verticalAlignment": "MIDDLE",
        }),
        # Row 34 — savings rate
        rep(33, 0, 34, 4, {
            "backgroundColor": gray_lt,
            "textFormat": {"fontSize": 10},
            "verticalAlignment": "MIDDLE",
        }),
        # Column widths: A=220, B=120, C=110, D=30(spacer), E=180, F=120
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 220}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 3},
            "properties": {"pixelSize": 120}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4},
            "properties": {"pixelSize": 30}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5},
            "properties": {"pixelSize": 180}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 6},
            "properties": {"pixelSize": 120}, "fields": "pixelSize"
        }},
        # Row heights
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 0, "endIndex": 2},
            "properties": {"pixelSize": 34}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 2, "endIndex": 35},
            "properties": {"pixelSize": 22}, "fields": "pixelSize"
        }},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "ROWS", "startIndex": 32, "endIndex": 33},
            "properties": {"pixelSize": 42}, "fields": "pixelSize"
        }},
        # Currency format B col (rows 6-34)
        {"repeatCell": {
            "range": rng(5, 1, 34, 2),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00;($#,##0.00);\"-\""},
                "horizontalAlignment": "RIGHT",
            }},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
        }},
        # Percent format C col (rows 6-34)
        {"repeatCell": {
            "range": rng(5, 2, 34, 3),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "0.0%"},
                "horizontalAlignment": "CENTER",
            }},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
        }},
        # Savings rate B34 — percent
        {"repeatCell": {
            "range": rng(33, 1, 34, 2),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "0.0%"},
                "horizontalAlignment": "RIGHT",
            }},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
        }},
        # Currency format F col (pie chart data)
        {"repeatCell": {
            "range": rng(0, 5, 35, 6),
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00"},
                "horizontalAlignment": "RIGHT",
            }},
            "fields": "userEnteredFormat(numberFormat,horizontalAlignment)"
        }},
        # Conditional: MONEY LEFT OVER — green text if positive
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [rng(32, 0, 33, 4)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_GREATER",
                                  "values": [{"userEnteredValue": "0"}]},
                    "format": {"backgroundColor": {"red":0.063,"green":0.373,"blue":0.2},
                               "textFormat": {"foregroundColor": white}}
                }
            }, "index": 0
        }},
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [rng(32, 0, 33, 4)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_LESS",
                                  "values": [{"userEnteredValue": "0"}]},
                    "format": {"backgroundColor": {"red":0.6,"green":0.1,"blue":0.1},
                               "textFormat": {"foregroundColor": white}}
                }
            }, "index": 1
        }},
        # Discretionary rows highlight if > 10% of income
        {"addConditionalFormatRule": {
            "rule": {
                "ranges": [rng(13, 2, 30, 3)],
                "booleanRule": {
                    "condition": {"type": "NUMBER_GREATER",
                                  "values": [{"userEnteredValue": "0.1"}]},
                    "format": {"backgroundColor": neg_bg}
                }
            }, "index": 2
        }},
        # Borders
        {"updateBorders": {
            "range": rng(0, 0, 34, 4),
            "innerHorizontal": {"style": "SOLID",
                                "color": {"red": 0.85, "green": 0.85, "blue": 0.85}},
            "innerVertical":   {"style": "SOLID",
                                "color": {"red": 0.85, "green": 0.85, "blue": 0.85}},
        }},
    ]

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
            print(f"  Format error: {r.stderr[:200]}")
            return False
        time.sleep(1)
    return True


def add_pie_chart():
    """Add a pie chart anchored below the data (row 37, col A), visible without scrolling right."""
    # E col = category names (col index 4), F col = amounts (col index 5)
    # Data rows 14-30 in sheet = 0-indexed 13-29
    chart_request = {
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Where Your Money Went",
                    "titleTextFormat": {
                        "bold": True,
                        "fontSize": 14,
                        "foregroundColor": {"red": 0.118, "green": 0.227, "blue": 0.373}
                    },
                    "pieChart": {
                        "legendPosition": "RIGHT_LEGEND",
                        "threeDimensional": False,
                        "domain": {
                            "sourceRange": {
                                "sources": [{
                                    "sheetId": DASHBOARD_SHEET_ID,
                                    "startRowIndex": 13,
                                    "endRowIndex": 30,
                                    "startColumnIndex": 4,
                                    "endColumnIndex": 5
                                }]
                            }
                        },
                        "series": {
                            "sourceRange": {
                                "sources": [{
                                    "sheetId": DASHBOARD_SHEET_ID,
                                    "startRowIndex": 13,
                                    "endRowIndex": 30,
                                    "startColumnIndex": 5,
                                    "endColumnIndex": 6
                                }]
                            }
                        }
                    },
                    "backgroundColor": {"red": 1, "green": 1, "blue": 1},
                    "fontName": "Arial"
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": DASHBOARD_SHEET_ID,
                            "rowIndex": 36,
                            "columnIndex": 0
                        },
                        "offsetXPixels": 0,
                        "offsetYPixels": 0,
                        "widthPixels": 600,
                        "heightPixels": 420
                    }
                }
            }
        }
    }

    r = subprocess.run(
        ["node", GWS_SCRIPT, "sheets", "spreadsheets", "batchUpdate",
         "--params", json.dumps({"spreadsheetId": SPREADSHEET_ID}),
         "--json",   json.dumps({"requests": [chart_request]}, ensure_ascii=False)],
        capture_output=True, text=True, encoding="utf-8"
    )
    return r.returncode == 0, r.stderr[:300]


def clear_dashboard():
    # Clear values
    subprocess.run(
        ["node", GWS_SCRIPT, "sheets", "spreadsheets", "values", "clear",
         "--params", json.dumps({
             "spreadsheetId": SPREADSHEET_ID,
             "range": chr(0x1F3E0) + " Dashboard!A1:L200"
         })],
        capture_output=True, text=True, encoding="utf-8"
    )
    # Unmerge all cells + delete existing charts
    r = subprocess.run(
        ["node", GWS_SCRIPT, "sheets", "spreadsheets", "get",
         "--params", json.dumps({"spreadsheetId": SPREADSHEET_ID})],
        capture_output=True, text=True, encoding="utf-8"
    )
    try:
        d = json.loads(r.stdout)
        batch_reqs = [
            # Unmerge everything so writes aren't blocked by old merged cells
            {"unmergeCells": {"range": {
                "sheetId": DASHBOARD_SHEET_ID,
                "startRowIndex": 0, "endRowIndex": 100,
                "startColumnIndex": 0, "endColumnIndex": 20
            }}}
        ]
        for sheet in d.get("sheets", []):
            if sheet["properties"]["sheetId"] == DASHBOARD_SHEET_ID:
                for chart in sheet.get("charts", []):
                    batch_reqs.append({"deleteEmbeddedObject": {"objectId": chart["chartId"]}})
        subprocess.run(
            ["node", GWS_SCRIPT, "sheets", "spreadsheets", "batchUpdate",
             "--params", json.dumps({"spreadsheetId": SPREADSHEET_ID}),
             "--json",   json.dumps({"requests": batch_reqs})],
            capture_output=True, text=True, encoding="utf-8"
        )
    except Exception:
        pass


def main():
    print("Setting up Dashboard...")

    print("  Clearing existing content...")
    clear_dashboard()
    time.sleep(2)

    print("  Writing data and formulas...")
    ok = write_values(ROWS)
    if not ok:
        print("  Aborted due to write error.")
        return
    time.sleep(2)

    print("  Applying formatting...")
    ok = apply_formatting()
    print(f"  Formatting: {'OK' if ok else 'ERROR'}")
    time.sleep(2)

    print("  Adding pie chart...")
    ok, err = add_pie_chart()
    if ok:
        print("  Pie chart: OK")
    else:
        print(f"  Pie chart: ERROR — {err}")

    print(f"\nDone.")
    print(f"  Change B2 (1-12) to switch months.")
    print(f"  https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()
