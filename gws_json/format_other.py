import json
import subprocess
import sys
import os

SPREADSHEET_ID = "1DEaFJvnXOM_B9GglT6sKXZzq0lCYF__EP_reMm3qu-w"
GWS = r"C:\Users\Bruce\AppData\Roaming\npm\gws.cmd"
OUTDIR = r"C:\Users\Bruce\source\repos\ClaudeBudget\gws_json\months"

def navy():
    return {"red": 0.118, "green": 0.227, "blue": 0.373}
def white():
    return {"red": 1, "green": 1, "blue": 1}
def med_blue():
    return {"red": 0.145, "green": 0.388, "blue": 0.922}
def light_blue():
    return {"red": 0.859, "green": 0.918, "blue": 0.996}
def total_row_bg():
    return {"red": 0.953, "green": 0.957, "blue": 0.965}


def run_batch(name, requests):
    payload = {"requests": requests}
    payload_file = os.path.join(OUTDIR, f"{name}_fmt.json")
    ps1_file = os.path.join(OUTDIR, f"{name}_fmt.ps1")

    with open(payload_file, "w", encoding="utf-8") as f:
        json.dump(payload, f)

    params_json = json.dumps({"spreadsheetId": SPREADSHEET_ID})
    ps_content = f"""$json = Get-Content -Path '{payload_file}' -Raw -Encoding UTF8
& '{GWS}' sheets spreadsheets batchUpdate --params '{params_json}' --json $json
"""
    with open(ps1_file, "w", encoding="utf-8") as f:
        f.write(ps_content)

    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1_file],
        capture_output=True, text=True, encoding="utf-8"
    )
    if result.returncode == 0:
        print(f"  {name}: OK")
    else:
        print(f"  {name}: ERROR {result.returncode} | {result.stderr[:200]} | {result.stdout[:200]}")
        raise RuntimeError(f"Failed: {name}")


def repeat_cell(sheet_id, r1, r2, c1, c2, bg=None, fg=None, bold=None, fontsize=None, halign=None, numfmt=None):
    fmt = {}
    if bg: fmt["backgroundColor"] = bg
    tf = {}
    if fg: tf["foregroundColor"] = fg
    if bold is not None: tf["bold"] = bold
    if fontsize: tf["fontSize"] = fontsize
    if tf: fmt["textFormat"] = tf
    if halign: fmt["horizontalAlignment"] = halign
    if numfmt: fmt["numberFormat"] = numfmt

    fields = []
    if bg: fields.append("backgroundColor")
    if tf: fields.append("textFormat")
    if halign: fields.append("horizontalAlignment")
    if numfmt: fields.append("numberFormat")

    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2, "startColumnIndex": c1, "endColumnIndex": c2},
            "cell": {"userEnteredFormat": fmt},
            "fields": "userEnteredFormat(" + ",".join(fields) + ")"
        }
    }

def merge(sheet_id, r1, r2, c1, c2):
    return {
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2, "startColumnIndex": c1, "endColumnIndex": c2},
            "mergeType": "MERGE_ALL"
        }
    }

def freeze(sheet_id, rows=0, cols=0):
    props = {}
    fields = []
    if rows: props["frozenRowCount"] = rows; fields.append("gridProperties.frozenRowCount")
    if cols: props["frozenColumnCount"] = cols; fields.append("gridProperties.frozenColumnCount")
    return {
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": props},
            "fields": ",".join(fields)
        }
    }

def currency_fmt(sheet_id, r1, r2, c1, c2):
    return repeat_cell(sheet_id, r1, r2, c1, c2, numfmt={"type": "CURRENCY", "pattern": "$#,##0.00"})

def pct_fmt(sheet_id, r1, r2, c1, c2):
    return repeat_cell(sheet_id, r1, r2, c1, c2, numfmt={"type": "PERCENT", "pattern": "0.0%"})


# ============================================================
# DEBT PAYOFF (698513475)
# ============================================================
DEBT = 698513475
debt_reqs = [
    merge(DEBT, 0, 1, 0, 9),
    repeat_cell(DEBT, 0, 1, 0, 9, bg=navy(), fg=white(), bold=True, fontsize=18, halign="CENTER"),
    repeat_cell(DEBT, 2, 3, 0, 9, bg=light_blue(), bold=True),
    repeat_cell(DEBT, 12, 13, 0, 9, bg=total_row_bg(), bold=True),
    merge(DEBT, 14, 15, 0, 8),
    repeat_cell(DEBT, 14, 15, 0, 8, bg=navy(), fg=white(), bold=True),
    repeat_cell(DEBT, 15, 16, 0, 8, bg=light_blue(), bold=True),
    repeat_cell(DEBT, 77, 78, 0, 8, bg=med_blue(), fg=white(), bold=True),
    freeze(DEBT, rows=3),
    currency_fmt(DEBT, 3, 12, 1, 4),
]
run_batch("debt", debt_reqs)


# ============================================================
# SINKING FUNDS (21742921)
# ============================================================
SINK = 21742921
sink_reqs = [
    merge(SINK, 0, 1, 0, 8),
    repeat_cell(SINK, 0, 1, 0, 8, bg=navy(), fg=white(), bold=True, fontsize=18, halign="CENTER"),
    repeat_cell(SINK, 1, 2, 0, 8, bg=light_blue(), bold=True),
    repeat_cell(SINK, 11, 12, 0, 8, bg=med_blue(), fg=white(), bold=True),
    repeat_cell(SINK, 16, 17, 0, 8, bg=med_blue(), fg=white(), bold=True),
    repeat_cell(SINK, 27, 28, 0, 8, bg=med_blue(), fg=white(), bold=True),
    freeze(SINK, rows=2),
    currency_fmt(SINK, 2, 11, 1, 5),
    pct_fmt(SINK, 2, 11, 5, 6),
]
run_batch("sinking", sink_reqs)


# ============================================================
# SUBSCRIPTIONS (1334531665)
# ============================================================
SUBS = 1334531665
subs_reqs = [
    merge(SUBS, 0, 1, 0, 12),
    repeat_cell(SUBS, 0, 1, 0, 12, bg=navy(), fg=white(), bold=True, fontsize=18, halign="CENTER"),
    repeat_cell(SUBS, 1, 2, 0, 12, bg=light_blue(), bold=True),
    repeat_cell(SUBS, 18, 19, 0, 12, bg=med_blue(), fg=white(), bold=True),
    repeat_cell(SUBS, 29, 30, 0, 12, bg=med_blue(), fg=white(), bold=True),
    freeze(SUBS, rows=2),
    currency_fmt(SUBS, 2, 18, 3, 6),
]
run_batch("subs", subs_reqs)


# ============================================================
# NET WORTH (903198120)
# ============================================================
NW = 903198120
nw_reqs = [
    merge(NW, 0, 1, 0, 4),
    repeat_cell(NW, 0, 1, 0, 4, bg=navy(), fg=white(), bold=True, fontsize=18, halign="CENTER"),
    repeat_cell(NW, 1, 2, 0, 4, bg=navy(), fg=white(), bold=True),
    repeat_cell(NW, 2, 3, 0, 4, bg=light_blue(), bold=True),
    repeat_cell(NW, 12, 13, 0, 4, bg=total_row_bg(), bold=True),
    repeat_cell(NW, 14, 15, 0, 4, bg=navy(), fg=white(), bold=True),
    repeat_cell(NW, 15, 16, 0, 4, bg=light_blue(), bold=True),
    repeat_cell(NW, 24, 25, 0, 4, bg=total_row_bg(), bold=True),
    repeat_cell(NW, 26, 27, 0, 4, bg=navy(), fg=white(), bold=True),
    repeat_cell(NW, 29, 30, 1, 2, bold=True, fontsize=18),
    repeat_cell(NW, 33, 34, 0, 4, bg=med_blue(), fg=white(), bold=True),
    freeze(NW, rows=3),
    currency_fmt(NW, 3, 35, 2, 4),
]
run_batch("networth", nw_reqs)


# ============================================================
# ANNUAL SUMMARY (503737128)
# ============================================================
ANN = 503737128
ann_reqs = [
    merge(ANN, 0, 1, 0, 7),
    repeat_cell(ANN, 0, 1, 0, 7, bg=navy(), fg=white(), bold=True, fontsize=18, halign="CENTER"),
    repeat_cell(ANN, 1, 2, 0, 7, bg=light_blue(), bold=True),
    repeat_cell(ANN, 14, 15, 0, 7, bg=total_row_bg(), bold=True),
    repeat_cell(ANN, 16, 17, 0, 7, bg=med_blue(), fg=white(), bold=True),
    repeat_cell(ANN, 24, 25, 0, 7, bg=med_blue(), fg=white(), bold=True),
    repeat_cell(ANN, 25, 26, 0, 7, bg=light_blue(), bold=True),
    freeze(ANN, rows=2),
    currency_fmt(ANN, 2, 15, 1, 6),
    pct_fmt(ANN, 2, 15, 6, 7),
]
run_batch("annual", ann_reqs)


# ============================================================
# SMART CALENDAR (512363891)
# ============================================================
CAL = 512363891
cal_reqs = [
    merge(CAL, 0, 1, 0, 7),
    repeat_cell(CAL, 0, 1, 0, 7, bg=navy(), fg=white(), bold=True, fontsize=18, halign="CENTER"),
    repeat_cell(CAL, 2, 3, 0, 7, bg=med_blue(), fg=white(), bold=True, halign="CENTER"),
    repeat_cell(CAL, 9, 10, 0, 7, bg=med_blue(), fg=white(), bold=True),
    repeat_cell(CAL, 10, 11, 0, 7, bg=light_blue(), bold=True),
    repeat_cell(CAL, 25, 26, 0, 7, bg=med_blue(), fg=white(), bold=True),
    freeze(CAL, rows=3),
    currency_fmt(CAL, 11, 25, 3, 4),
]
run_batch("calendar", cal_reqs)

print("All other sheets formatted!")
