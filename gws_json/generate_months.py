import json
import subprocess
import sys
import os

SPREADSHEET_ID = "1DEaFJvnXOM_B9GglT6sKXZzq0lCYF__EP_reMm3qu-w"
GWS = r"C:\Users\Bruce\AppData\Roaming\npm\gws.cmd"
OUTDIR = r"C:\Users\Bruce\source\repos\ClaudeBudget\gws_json\months"
os.makedirs(OUTDIR, exist_ok=True)

months = [
    ("Mar", "MARCH", "2025-03-01", "2025-03-31", 31),
    ("Apr", "APRIL", "2025-04-01", "2025-04-30", 30),
    ("May", "MAY", "2025-05-01", "2025-05-31", 31),
    ("Jun", "JUNE", "2025-06-01", "2025-06-30", 30),
    ("Jul", "JULY", "2025-07-01", "2025-07-31", 31),
    ("Aug", "AUGUST", "2025-08-01", "2025-08-31", 31),
    ("Sep", "SEPTEMBER", "2025-09-01", "2025-09-30", 30),
    ("Oct", "OCTOBER", "2025-10-01", "2025-10-31", 31),
    ("Nov", "NOVEMBER", "2025-11-01", "2025-11-30", 30),
    ("Dec", "DECEMBER", "2025-12-01", "2025-12-31", 31),
]

def make_month_data(abbr, name_upper, start, end, days):
    tab = f"📅 {abbr}"
    d = []

    def r(rng, vals):
        d.append({"range": f"'{tab}'!{rng}", "values": vals})

    r("A1", [[f"{name_upper} 2025"]])
    r("A2", [["50/30/20 Budget Dashboard"]])
    r("A4:F4", [["Start Date", start, "End Date", end, "Days", days]])
    r("A6", [["INCOME"]])
    r("A7:E7", [["Category", "Source Name", "Expected $", "Actual $", "Variance"]])
    r("A8:E8", [["Primary", "Paycheck 1", 4500, 0, "=D8-C8"]])
    r("A9:E9", [["Primary", "Paycheck 2", 4500, 0, "=D9-C9"]])
    r("A10:E10", [["Freelance", "Freelance Project", 1200, 0, "=D10-C10"]])
    r("A11:E11", [["Side Hustle", "Etsy/Other", 300, 0, "=D11-C11"]])
    r("A12:E12", [["Dividend", "Dividend Income", 0, 0, "=D12-C12"]])
    r("A13:E13", [["Other", "Other Income 1", 0, 0, "=D13-C13"]])
    r("A14:E14", [["Other", "Other Income 2", 0, 0, "=D14-C14"]])
    r("A16:E16", [["TOTAL INCOME", "", "=SUM(C8:C14)", "=SUM(D8:D14)", "=D16-C16"]])
    r("A18", [["🏠 NEEDS — 50%"]])
    r("A19:F19", [["Category", "Due Date", "Expected $", "Actual $", "Action", "Progress %"]])

    needs = [
        ("Housing/Rent", 1800), ("Electricity", 120), ("Water", 60),
        ("Internet", 65), ("Mobile Phone", 85), ("Car Payment", 485),
        ("Car Insurance", 175), ("Fuel/Transportation", 200), ("Groceries", 600),
        ("Health Insurance", 320), ("Medical", 50), ("Min Debt Payments", 578),
        ("Childcare", 0), ("Other Necessity 1", 0)
    ]
    for i, (cat, amt) in enumerate(needs):
        row = 20 + i
        r(f"A{row}:F{row}", [[cat, "", amt, 0, f'=IF(D{row}>=C{row},"✅","⬜")', f"=IFERROR(D{row}/C{row},0)"]])

    r("A34:F34", [["TOTAL NEEDS", "", "=SUM(C20:C33)", "=SUM(D20:D33)", "", "=IFERROR(D34/C34,0)"]])
    r("A36", [["🎯 WANTS — 30%"]])
    r("A37:F37", [["Category", "Due Date", "Expected $", "Actual $", "Action", "Progress %"]])

    wants = [
        ("Dining Out", 300), ("Entertainment", 150), ("Streaming Services", 65),
        ("Shopping", 300), ("Hobbies", 100), ("Gym/Fitness", 50),
        ("Beauty", 60), ("Travel", 0), ("Amazon/Online", 100),
        ("Gifts", 50), ("Non-essential Subs", 30), ("Other Want 1", 0)
    ]
    for i, (cat, amt) in enumerate(wants):
        row = 38 + i
        r(f"A{row}:F{row}", [[cat, "", amt, 0, f'=IF(D{row}>=C{row},"✅","⬜")', f"=IFERROR(D{row}/C{row},0)"]])

    r("A50:F50", [["TOTAL WANTS", "", "=SUM(C38:C49)", "=SUM(D38:D49)", "", "=IFERROR(D50/C50,0)"]])
    r("A52", [["💰 SAVINGS & DEBT — 20%"]])
    r("A53:F53", [["Category", "Due Date", "Expected $", "Actual $", "Action", "Progress %"]])

    savings = [
        ("Emergency Fund", 500), ("Vacation Fund", 300), ("Car Fund", 200),
        ("Retirement/401k", 450), ("Investment Account", 0),
        ("Extra CC Payment 1", 200), ("Extra CC Payment 2", 100),
        ("Other Savings Goal 1", 100)
    ]
    for i, (cat, amt) in enumerate(savings):
        row = 54 + i
        r(f"A{row}:F{row}", [[cat, "", amt, 0, f'=IF(D{row}>=C{row},"✅","⬜")', f"=IFERROR(D{row}/C{row},0)"]])

    r("A62:F62", [["TOTAL SAVINGS/DEBT", "", "=SUM(C54:C61)", "=SUM(D54:D61)", "", "=IFERROR(D62/C62,0)"]])
    r("A64", [["TOTALS SUMMARY"]])
    r("A65:E65", [["", "Target $", "Actual $", "Remaining", "% of Income"]])
    r("A66:E66", [["Total Needs", "=D16*'🏠 Dashboard'!B6", "=D34", "=B66-C66", "=IFERROR(C66/D16,0)"]])
    r("A67:E67", [["Total Wants", "=D16*'🏠 Dashboard'!B7", "=D50", "=B67-C67", "=IFERROR(C67/D16,0)"]])
    r("A68:E68", [["Total Savings/Debt", "=D16*'🏠 Dashboard'!B8", "=D62", "=B68-C68", "=IFERROR(C68/D16,0)"]])
    r("A69:C69", [["GRAND TOTAL EXPENSES", "", "=C66+C67+C68"]])
    r("A70:C70", [["NET CASH FLOW", "", "=D16-C69"]])
    r("A71:C71", [["Rollover to Next Month", "", "=C70"]])

    return {"valueInputOption": "USER_ENTERED", "data": d}


for abbr, name_upper, start, end, days in months:
    print(f"Uploading {name_upper}...", end=" ", flush=True)
    payload = make_month_data(abbr, name_upper, start, end, days)

    # Write payload and a PowerShell script to separate files
    payload_file = os.path.join(OUTDIR, f"{abbr}_payload.json")
    ps1_file = os.path.join(OUTDIR, f"{abbr}_upload.ps1")

    with open(payload_file, "w", encoding="utf-8") as f:
        json.dump(payload, f)

    params_json = json.dumps({"spreadsheetId": SPREADSHEET_ID})
    # Escape single quotes for PowerShell
    params_ps = params_json.replace("'", "''")

    ps_content = f"""$json = Get-Content -Path '{payload_file}' -Raw -Encoding UTF8
& '{GWS}' sheets spreadsheets values batchUpdate --params '{params_ps}' --json $json
"""
    with open(ps1_file, "w", encoding="utf-8") as f:
        f.write(ps_content)

    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1_file],
        capture_output=True, text=True, encoding="utf-8"
    )

    if result.returncode == 0:
        out = result.stdout
        lines = [l for l in out.split("\n") if l.strip() and not l.startswith("Using")]
        try:
            data = json.loads("\n".join(lines))
            print(f"OK ({data.get('totalUpdatedCells', '?')} cells)")
        except Exception as e:
            print(f"OK (output: {out[:200]})")
    else:
        print(f"ERROR rc={result.returncode}: {result.stderr[:300]} | stdout: {result.stdout[:200]}")
        sys.exit(1)

print("All months done!")
