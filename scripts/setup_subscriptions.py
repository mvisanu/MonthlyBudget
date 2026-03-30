#!/usr/bin/env python3
"""
Rewrites the Subscriptions tab with verified 2026 subscriptions from bank data,
plus a historical section for old/unconfirmed items.

Usage:
    python scripts/setup_subscriptions.py
"""

import io, json, subprocess, sys, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SPREADSHEET_ID = "1DEaFJvnXOM_B9GglT6sKXZzq0lCYF__EP_reMm3qu-w"
SUBS_SHEET_ID  = 1334531665
GWS_SCRIPT     = r"C:\Users\Bruce\AppData\Roaming\npm\node_modules\@googleworkspace\cli\run-gws.js"

# ── Layout ────────────────────────────────────────────────────────────────────
# Cols: A=Service, B=Category, C=Billing Cycle, D=Amount/Cycle, E=Monthly Cost, F=Status/Notes
#
# SECTION 1: ACTIVE 2026 (confirmed from bank transactions)
# SECTION 2: HISTORICAL / UNCONFIRMED (old placeholders — may be on credit card)
# ─────────────────────────────────────────────────────────────────────────────

ROWS = [
    # Row 1 — Title
    ["SUBSCRIPTION TRACKER 2026", "", "", "", "", ""],
    # Row 2 — column headers
    ["Service", "Category", "Billing Cycle", "Amount/Cycle ($)", "Est. Monthly ($)", "Status / Notes"],

    # ── ACTIVE 2026 ──────────────────────────────────────────────────────────
    ["ACTIVE — Confirmed from Bank (2026)", "", "", "", "", ""],
    ["Claude.ai Pro",       "AI / Software", "Annual",  "89.15",  "=D5/12",  "Active — charged Mar 2026"],
    ["ChatGPT (OpenAI)",    "AI / Software", "Monthly", "20.00",  "20.00",   "Active — charged Feb 2026"],
    ["YouTube Premium",     "Streaming",     "Monthly", "15.88",  "15.88",   "Active — charged Feb 2026"],
    ["Rocket Money Premium","Finance",       "Monthly", "11.00",  "11.00",   "Active — charged Jan 2026"],
    ["Tonal",               "Fitness",       "Monthly", "66.15",  "66.15",   "Active — charged Mar 2026"],
    # Row 9 — subtotal
    ["ACTIVE TOTAL / MONTH", "", "", "", "=SUM(E5:E8)", ""],

    # blank
    ["", "", "", "", "", ""],

    # ── CANCELLED / NO LONGER SUBSCRIBED ─────────────────────────────────────
    ["CANCELLED — No Longer Subscribed", "", "", "", "", ""],
    # (user will fill these in — placeholder rows)
    ["", "", "", "", "", ""],
    ["", "", "", "", "", ""],

    # blank
    ["", "", "", "", "", ""],

    # ── HISTORICAL / UNCONFIRMED ──────────────────────────────────────────────
    ["HISTORICAL / UNCONFIRMED — Not seen in checking account", "", "", "", "", ""],
    ["(These may be charged to a credit card — verify each one)", "", "", "", "", ""],
    ["Netflix",            "Streaming", "Monthly", "15.49", "15.49", "Unconfirmed — not in bank transactions"],
    ["Spotify",            "Streaming", "Monthly", "9.99",  "9.99",  "Unconfirmed — not in bank transactions"],
    ["Disney+",            "Streaming", "Monthly", "13.99", "13.99", "Unconfirmed — not in bank transactions"],
    ["Adobe CC",           "Software",  "Annual",  "599.88","49.99", "Unconfirmed — not in bank transactions"],
    ["Microsoft 365",      "Software",  "Annual",  "99.99", "8.33",  "Unconfirmed — not in bank transactions"],
    ["Notion",             "Software",  "Annual",  "96.00", "8.00",  "Unconfirmed — not in bank transactions"],
    ["LastPass",           "Software",  "Annual",  "36.00", "3.00",  "Unconfirmed — not in bank transactions"],
    ["Life Insurance",     "Insurance", "Monthly", "45.00", "45.00", "Unconfirmed — not in bank transactions"],
    ["Renters Insurance",  "Insurance", "Annual",  "180.00","15.00", "Unconfirmed — not in bank transactions"],
    ["Credit Monitoring",  "Finance",   "Monthly", "19.99", "19.99", "Unconfirmed — not in bank transactions"],
    ["Gym Membership",     "Health",    "Monthly", "50.00", "50.00", "Unconfirmed — not in bank transactions"],
    ["Meditation App",     "Health",    "Annual",  "69.99", "5.83",  "Unconfirmed — not in bank transactions"],
    ["Amazon Prime",       "Other",     "Annual",  "139.00","11.58", "Unconfirmed — not in bank transactions"],
]

navy    = {"red": 0.118, "green": 0.227, "blue": 0.373}
white   = {"red": 1.0,   "green": 1.0,   "blue": 1.0}
green   = {"red": 0.180, "green": 0.490, "blue": 0.196}
red_lt  = {"red": 0.996, "green": 0.886, "blue": 0.886}
grn_lt  = {"red": 0.851, "green": 0.953, "blue": 0.867}
yel_lt  = {"red": 1.000, "green": 0.973, "blue": 0.820}
gray_lt = {"red": 0.953, "green": 0.957, "blue": 0.965}
gray_md = {"red": 0.878, "green": 0.878, "blue": 0.878}
alt     = {"red": 0.976, "green": 0.980, "blue": 0.984}
red_hdr = {"red": 0.698, "green": 0.133, "blue": 0.133}
yel_hdr = {"red": 0.737, "green": 0.604, "blue": 0.118}


def gws_call(cmd_args, payload):
    r = subprocess.run(
        ["node", GWS_SCRIPT] + cmd_args +
        ["--params", json.dumps({"spreadsheetId": SPREADSHEET_ID}),
         "--json",   json.dumps(payload, ensure_ascii=False)],
        capture_output=True, text=True, encoding="utf-8"
    )
    return r.returncode == 0, r.stderr[:300] if r.returncode != 0 else ""


def rng(r1, c1, r2, c2):
    return {"sheetId": SUBS_SHEET_ID, "startRowIndex": r1, "endRowIndex": r2,
            "startColumnIndex": c1, "endColumnIndex": c2}


def rep(r1, c1, r2, c2, fmt, fields=None):
    return {"repeatCell": {
        "range": rng(r1, c1, r2, c2),
        "cell": {"userEnteredFormat": fmt},
        "fields": fields or "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,numberFormat)"
    }}


def write_values():
    ok, err = gws_call(
        ["sheets", "spreadsheets", "values", "clear"],
        {"range": chr(0x1F501) + " Subscriptions!A1:F60"}
    )

    start = 1
    for i in range(0, len(ROWS), 5):
        chunk = ROWS[i:i+5]
        payload = {
            "valueInputOption": "USER_ENTERED",
            "data": [{"range": chr(0x1F501) + f" Subscriptions!A{start}", "values": chunk}]
        }
        ok, err = gws_call(["sheets", "spreadsheets", "values", "batchUpdate"], payload)
        if not ok:
            print(f"  Write error at row {start}: {err}")
            return False
        start += len(chunk)
        time.sleep(0.3)
    return True


def apply_formatting():
    def hdr(r, bg, size=11):
        return rep(r, 0, r+1, 6, {
            "backgroundColor": bg,
            "textFormat": {"bold": True, "fontSize": size, "foregroundColor": white},
            "horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE",
        })

    requests = [
        # Column widths
        {"updateDimensionProperties": {"range": {"sheetId": SUBS_SHEET_ID, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 200}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": SUBS_SHEET_ID, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 120}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": SUBS_SHEET_ID, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 110}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": SUBS_SHEET_ID, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 130}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": SUBS_SHEET_ID, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5}, "properties": {"pixelSize": 130}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": SUBS_SHEET_ID, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 6}, "properties": {"pixelSize": 300}, "fields": "pixelSize"}},
        # Row 1 — title
        hdr(0, navy, 14),
        # Row 2 — col headers
        rep(1, 0, 2, 6, {
            "backgroundColor": {"red": 0.267, "green": 0.431, "blue": 0.643},
            "textFormat": {"bold": True, "fontSize": 10, "foregroundColor": white},
            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
        }),
        # Row 3 — ACTIVE section header (green)
        hdr(2, green),
        # Rows 4-8 — active subscriptions (alternating green tint)
        *[rep(r, 0, r+1, 6, {
            "backgroundColor": grn_lt if i % 2 == 0 else white,
            "textFormat": {"fontSize": 10},
            "verticalAlignment": "MIDDLE",
        }) for i, r in enumerate(range(3, 8))],
        # Row 9 — active subtotal
        rep(8, 0, 9, 6, {
            "backgroundColor": gray_md,
            "textFormat": {"bold": True, "fontSize": 10},
            "verticalAlignment": "MIDDLE",
        }),
        # Row 11 (index 10) — blank, Row 12 (index 11) — CANCELLED header (red)
        hdr(11, red_hdr),
        # Rows 13-14 — cancelled placeholder rows
        rep(12, 0, 14, 6, {
            "backgroundColor": red_lt,
            "textFormat": {"fontSize": 10, "italic": True, "foregroundColor": {"red": 0.5, "green": 0.0, "blue": 0.0}},
            "verticalAlignment": "MIDDLE",
        }),
        # Row 16 (index 15) — blank, Row 17 (index 16) — HISTORICAL header (yellow/amber)
        hdr(16, yel_hdr),
        # Row 17 — note row
        rep(17, 0, 18, 6, {
            "backgroundColor": yel_lt,
            "textFormat": {"fontSize": 9, "italic": True},
            "verticalAlignment": "MIDDLE",
        }),
        # Rows 18-30 — historical items (alternating yellow tint)
        *[rep(r, 0, r+1, 6, {
            "backgroundColor": yel_lt if i % 2 == 0 else white,
            "textFormat": {"fontSize": 10, "foregroundColor": {"red": 0.4, "green": 0.4, "blue": 0.4}},
            "verticalAlignment": "MIDDLE",
        }) for i, r in enumerate(range(18, 31))],
        # Currency format D and E cols
        rep(3, 3, 31, 5, {
            "numberFormat": {"type": "NUMBER", "pattern": "$#,##0.00"},
            "horizontalAlignment": "RIGHT",
        }, fields="userEnteredFormat(numberFormat,horizontalAlignment)"),
        # Freeze row 1
        {"updateSheetProperties": {
            "properties": {"sheetId": SUBS_SHEET_ID, "gridProperties": {"frozenRowCount": 2}},
            "fields": "gridProperties.frozenRowCount"
        }},
    ]

    for i in range(0, len(requests), 8):
        chunk = requests[i:i+8]
        ok, err = gws_call(["sheets", "spreadsheets", "batchUpdate"], {"requests": chunk})
        if not ok:
            print(f"  Format error: {err}")
            return False
        time.sleep(1)
    return True


def main():
    print("Updating Subscriptions tab...")
    print("  Writing data...", end=" ", flush=True)
    if not write_values():
        return
    print("OK")

    print("  Applying formatting...", end=" ", flush=True)
    if apply_formatting():
        print("OK")
    else:
        print("ERROR")

    print(f"\nDone.")
    print(f"  https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()
