import json
import subprocess
import sys
import os

SPREADSHEET_ID = "1DEaFJvnXOM_B9GglT6sKXZzq0lCYF__EP_reMm3qu-w"
GWS = r"C:\Users\Bruce\AppData\Roaming\npm\gws.cmd"
OUTDIR = r"C:\Users\Bruce\source\repos\ClaudeBudget\gws_json\months"

month_sheet_ids = {
    "Jan": 122166217,
    "Feb": 636285504,
    "Mar": 1742587835,
    "Apr": 278726061,
    "May": 1054479217,
    "Jun": 1278089603,
    "Jul": 542569607,
    "Aug": 555749728,
    "Sep": 1955324181,
    "Oct": 679307036,
    "Nov": 1270203974,
    "Dec": 1885490765,
}

def navy():
    return {"red": 0.118, "green": 0.227, "blue": 0.373}

def white():
    return {"red": 1, "green": 1, "blue": 1}

def med_blue():
    return {"red": 0.145, "green": 0.388, "blue": 0.922}

def light_blue():
    return {"red": 0.859, "green": 0.918, "blue": 0.996}

def total_row():
    return {"red": 0.953, "green": 0.957, "blue": 0.965}

def alt_row():
    return {"red": 0.976, "green": 0.980, "blue": 0.984}

def make_format_requests(sheet_id):
    requests = []

    def navy_header(start_row, end_row):
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row, "startColumnIndex": 0, "endColumnIndex": 6},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": navy(),
                        "textFormat": {"foregroundColor": white(), "bold": True}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        })

    def med_blue_header(start_row, end_row):
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row, "startColumnIndex": 0, "endColumnIndex": 6},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": med_blue(),
                        "textFormat": {"foregroundColor": white(), "bold": True}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        })

    def col_header(start_row, end_row):
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row, "startColumnIndex": 0, "endColumnIndex": 6},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": light_blue(),
                        "textFormat": {"bold": True}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        })

    def total_row_fmt(start_row, end_row):
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": start_row, "endRowIndex": end_row, "startColumnIndex": 0, "endColumnIndex": 6},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": total_row(),
                        "textFormat": {"bold": True}
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        })

    # Row 1: title - merge and navy
    requests.append({
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 6},
            "mergeType": "MERGE_ALL"
        }
    })
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 6},
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": navy(),
                    "textFormat": {"foregroundColor": white(), "fontSize": 18, "bold": True},
                    "horizontalAlignment": "CENTER"
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
        }
    })

    # Row 2: subtitle - merge and med blue
    requests.append({
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 6},
            "mergeType": "MERGE_ALL"
        }
    })
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 6},
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": med_blue(),
                    "textFormat": {"foregroundColor": white(), "fontSize": 12},
                    "horizontalAlignment": "CENTER"
                }
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
        }
    })

    # Row 6: INCOME section header - navy (row index 5)
    requests.append({
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": 5, "endRowIndex": 6, "startColumnIndex": 0, "endColumnIndex": 6},
            "mergeType": "MERGE_ALL"
        }
    })
    navy_header(5, 6)

    # Row 7: column headers (row index 6)
    col_header(6, 7)

    # Row 16: total income (row index 15)
    total_row_fmt(15, 16)

    # Row 18: NEEDS section header (row index 17)
    requests.append({
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": 17, "endRowIndex": 18, "startColumnIndex": 0, "endColumnIndex": 6},
            "mergeType": "MERGE_ALL"
        }
    })
    navy_header(17, 18)

    # Row 19: needs column headers (row index 18)
    col_header(18, 19)

    # Row 34: TOTAL NEEDS (row index 33)
    total_row_fmt(33, 34)

    # Row 36: WANTS section header (row index 35)
    requests.append({
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": 35, "endRowIndex": 36, "startColumnIndex": 0, "endColumnIndex": 6},
            "mergeType": "MERGE_ALL"
        }
    })
    navy_header(35, 36)

    # Row 37: wants column headers (row index 36)
    col_header(36, 37)

    # Row 50: TOTAL WANTS (row index 49)
    total_row_fmt(49, 50)

    # Row 52: SAVINGS section header (row index 51)
    requests.append({
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": 51, "endRowIndex": 52, "startColumnIndex": 0, "endColumnIndex": 6},
            "mergeType": "MERGE_ALL"
        }
    })
    navy_header(51, 52)

    # Row 53: savings column headers (row index 52)
    col_header(52, 53)

    # Row 62: TOTAL SAVINGS (row index 61)
    total_row_fmt(61, 62)

    # Row 64: TOTALS SUMMARY section header (row index 63)
    requests.append({
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": 63, "endRowIndex": 64, "startColumnIndex": 0, "endColumnIndex": 6},
            "mergeType": "MERGE_ALL"
        }
    })
    navy_header(63, 64)

    # Row 65: totals column headers (row index 64)
    col_header(64, 65)

    # Rows 69-71: grand total, net cash flow, rollover (row indices 68-70)
    total_row_fmt(68, 71)

    # Freeze row 3 and column A
    requests.append({
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": 3, "frozenColumnCount": 1}
            },
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
        }
    })

    # Currency format for Expected $ and Actual $ columns (C and D = cols 2,3)
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 7, "endRowIndex": 72, "startColumnIndex": 2, "endColumnIndex": 4},
            "cell": {
                "userEnteredFormat": {
                    "numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}
                }
            },
            "fields": "userEnteredFormat.numberFormat"
        }
    })

    # Percent format for Progress % column (F = col 5)
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 18, "endRowIndex": 72, "startColumnIndex": 5, "endColumnIndex": 6},
            "cell": {
                "userEnteredFormat": {
                    "numberFormat": {"type": "PERCENT", "pattern": "0.0%"}
                }
            },
            "fields": "userEnteredFormat.numberFormat"
        }
    })

    # % of income column E in totals section (rows 65-71, col 4)
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 64, "endRowIndex": 71, "startColumnIndex": 4, "endColumnIndex": 5},
            "cell": {
                "userEnteredFormat": {
                    "numberFormat": {"type": "PERCENT", "pattern": "0.0%"}
                }
            },
            "fields": "userEnteredFormat.numberFormat"
        }
    })

    return {"requests": requests}


for abbr, sheet_id in month_sheet_ids.items():
    print(f"Formatting {abbr}...", end=" ", flush=True)
    payload = make_format_requests(sheet_id)

    payload_file = os.path.join(OUTDIR, f"{abbr}_format.json")
    ps1_file = os.path.join(OUTDIR, f"{abbr}_format.ps1")

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
        print(f"OK")
    else:
        print(f"ERROR rc={result.returncode}: {result.stderr[:300]} | stdout: {result.stdout[:300]}")
        sys.exit(1)

print("All month formatting done!")
