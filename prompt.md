# Claude Code Prompt: Personal Finance Command Center (Google Sheets + Excel)

## OBJECTIVE

Build a fully-featured **Personal Finance Command Center** spreadsheet using the Google Workspace CLI (gsheet/gdrive tools). After completing the Google Sheets version, also export an `.xlsx` copy. The workbook must be production-ready — clean formulas with zero errors, dynamic cross-sheet references, professional minimal styling (white + gray + blue palette), and ready for real daily use.

---

## TOOLING NOTES

- Use `gsheet` / Google Workspace MCP CLI to create and populate the Google Sheet.
- Use Python + `openpyxl` to produce the `.xlsx` export.
- Run `scripts/recalc.py` on the `.xlsx` to validate zero formula errors before delivering.
- All formula values must be calculated by spreadsheet formulas — never hardcode computed values in Python.

---

## WORKBOOK STRUCTURE (10 Sheets)

Create sheets in this tab order:

| # | Tab Name | Description |
|---|----------|-------------|
| 1 | `🏠 Dashboard` | 50/30/20 Budget overview for current month |
| 2 | `📅 Jan` through `📅 Dec` | One monthly budget tab per month (12 tabs) |
| 3 | `💳 Debt Payoff` | Hybrid snowball/avalanche calculator |
| 4 | `🪣 Sinking Funds` | Goal-based savings buckets tracker |
| 5 | `📆 Smart Calendar` | Monthly bill/payday/subscription calendar |
| 6 | `🔁 Subscriptions` | Subscription tracker with renewal alerts |
| 7 | `📈 Net Worth` | Asset vs liability tracker |
| 8 | `📊 Annual Summary` | Year-over-year overview with charts |

> Note: Monthly tabs Jan–Dec can be identical in structure. Populate January with sample data; leave Feb–Dec with formulas only (no sample data).

---

## SHEET 1: 🏠 Dashboard

### Purpose
A single-glance monthly budget command center pulling live data from the active monthly tab.

### Layout & Content

**Section A — Header (Row 1–3)**
- Title: "Personal Finance Command Center"
- Subtitle: Current Month + Year (formula: `=TEXT(TODAY(),"MMMM YYYY")`)
- Input cell (yellow background): `[Month Selector]` — user types 1–12 to select the active month. Named range: `ActiveMonth`

**Section B — Budget Ratio Control Panel (Rows 5–9)**
- Label: "Budget Allocation Targets"
- Three input cells (blue text, yellow background — user edits these):
  - `Needs %` (default: 50%)
  - `Wants %` (default: 30%)
  - `Savings & Debt %` (default: 20%)
- Validation: Show warning if the three don't sum to 100%
  - Formula: `=IF(SUM(NeedsTarget,WantsTarget,SavingsTarget)<>1,"⚠️ Targets must total 100%","")`

**Section C — Income Summary (Rows 11–20)**
- Pull from active monthly tab using `INDIRECT()` referencing `ActiveMonth`
- Rows:
  - Income Source 1–5 (label + expected + actual)
  - Freelance/Side Hustle 1–3 (separate rows, same structure)
  - **Total Expected Income** (SUM formula)
  - **Total Actual Income** (SUM formula)
  - **Variance** (Actual − Expected)

**Section D — 50/30/20 Allocation Summary (Rows 22–30)**
- Three category cards side-by-side: Needs | Wants | Savings & Debt
- For each card show:
  - **Target $** = Total Actual Income × Target %
  - **Actual Spent $** = pulled from monthly tab
  - **Remaining $** = Target − Actual
  - **% Used** = Actual / Target (progress bar via conditional formatting)
- Color coding:
  - Under budget: blue fill (#DBEAFE)
  - Over budget: red fill (#FEE2E2)
  - Within 5%: gray fill (#F3F4F6)

**Section E — Cash Flow Summary (Rows 32–38)**
- Total Income − Total Expenses = **Net Cash Flow**
- Running monthly net (pull from each monthly tab via INDIRECT)
- Label: "Amount Left to Spend This Month"

**Section F — Key Metrics (Rows 40–46)**
- Savings Rate % = Savings / Total Income
- Debt Payoff Progress (% of total debt remaining, from Debt Payoff tab)
- Subscription Spend (total monthly, from Subscriptions tab)
- Net Worth (from Net Worth tab)
- Month-over-month income change %

---

## SHEETS 2–13: 📅 Monthly Tabs (Jan–Dec)

Each monthly tab is identical in structure. January gets sample data; Feb–Dec get formulas only.

### Layout

**Row 1–2: Header**
- Month name (hardcoded per tab, e.g., "JANUARY 2025")
- Sub-label: "50/30/20 Budget Dashboard"

**Rows 4–6: Period**
- Start Date | End Date | Days in Month

**Section: INCOME (Rows 8–22)**

| Column | Content |
|--------|---------|
| A | Category label |
| B | Income Source Name (user edits) |
| C | Expected $ (blue text — user input) |
| D | Actual $ (blue text — user input) |
| E | Variance (formula: D−C) |

- Rows for: Paycheck 1, Paycheck 2, Freelance, Side Hustle 1, Side Hustle 2, Dividend Income, Other Income 1–3
- **TOTAL INCOME row** at bottom (SUM formulas)

**Section: NEEDS — 50% (Rows 24–50)**

Pre-populated categories (user can rename):
- Housing/Rent, Electricity, Water, Gas, Internet, Mobile, Car Payment, Car Insurance, Fuel/Transportation, Groceries, Health Insurance, Medical, Minimum Debt Payments, Childcare, Other Necessity 1–3

Columns: Category | Due Date | Expected | Actual | Action | Progress %
- Action column: `=IF(Actual>=Expected,"✅","⬜")`
- Progress %: `=Actual/Expected` formatted as percentage

**Section: WANTS — 30% (Rows 52–75)**

Pre-populated categories:
- Dining Out, Entertainment, Streaming Services, Shopping, Hobbies, Gym/Fitness, Beauty, Travel, Amazon/Online, Gifts, Subscriptions (non-essential), Other Want 1–3

Same columns as Needs.

**Section: SAVINGS & DEBT — 20% (Rows 77–90)**

Pre-populated categories:
- Emergency Fund, Vacation Fund, Car Fund, Retirement/401k, Investment Account, Extra Debt Payment 1–3, Other Savings Goal 1–2

Same columns.

**Section: TOTALS (Rows 92–100)**
- Total Needs Actual | Target | Remaining | % of Income
- Total Wants Actual | Target | Remaining | % of Income  
- Total Savings/Debt Actual | Target | Remaining | % of Income
- Grand Total Expenses
- **Net Cash Flow** = Total Income − Grand Total Expenses
- **Rollover** = carry to next month (reference cell)

---

## SHEET: 💳 Debt Payoff Calculator

### Purpose
Hybrid snowball/avalanche payoff planner with toggle between methods.

### Layout

**Section A — Debt Input Table (Rows 3–20)**

Columns:
| Col | Header |
|-----|--------|
| A | Debt Name |
| B | Current Balance ($) — blue text, user input |
| C | Interest Rate (APR %) — blue text, user input |
| D | Minimum Payment ($) — blue text, user input |
| E | Snowball Order (auto-ranked by balance: `=RANK(B, $B$4:$B$20, 1)`) |
| F | Avalanche Order (auto-ranked by rate: `=RANK(C, $C$4:$C$20, 0)`) |
| G | Active Order (formula: `=IF(Method="Snowball", E, F)`) |
| H | Payoff Month (calculated from amortization below) |
| I | Total Interest Paid (calculated) |

Pre-populated with 8 sample debts: Credit Card 1–3, Auto Loan 1–2, Personal Loan, Student Loan, Medical Bill.

**Section B — Method Toggle (Row 1)**
- Dropdown cell (Data Validation): `Snowball | Avalanche | Hybrid`
- Named range: `Method`
- Hybrid = sort by balance first, then avalanche for ties

**Section C — Extra Payment Input**
- Cell: "Extra Monthly Payment Available $" (blue text, user input)
- Named range: `ExtraPayment`

**Section D — Payoff Timeline Table (Rows 22–80)**
- Month-by-month amortization for the #1 priority debt
- Columns: Month # | Month/Year | Starting Balance | Payment | Principal | Interest | Ending Balance | Notes
- Formulas use `IF(EndingBalance<=0, 0, ...)` to stop at payoff
- "Notes" column auto-populates "PAID OFF ✅" when balance hits 0, and "➡ Roll to next debt" for the extra payment rollover

**Section E — Summary Metrics (Rows 82–90)**
- Total Debt: `=SUM(B4:B20)`
- Months to Debt Free (Snowball method)
- Months to Debt Free (Avalanche method)
- Total Interest — Snowball
- Total Interest — Avalanche
- Interest Savings by choosing Avalanche
- Estimated Debt-Free Date

**Section F — Progress Bar**
- Visual: `=REPT("█", INT(PaidPct*20)) & REPT("░", 20-INT(PaidPct*20))`
- % Paid Off label

---

## SHEET: 🪣 Sinking Funds Tracker

### Purpose
Track 6–8 savings goal buckets with monthly contribution plan and progress.

### Layout

**Section A — Fund Table (Rows 3–20)**

8 pre-built fund buckets:
1. Emergency Fund (target: 3–6 months expenses)
2. Car Maintenance / Replacement
3. Vacation / Travel
4. Home Repair
5. Medical / HSA
6. Annual Subscriptions (pre-pay bucket)
7. Holiday / Gifts
8. [Custom — user-named]

Columns per fund:
| Col | Header |
|-----|---------|
| A | Fund Name |
| B | Goal Amount ($) — blue input |
| C | Target Date — blue input |
| D | Current Balance ($) — blue input |
| E | Monthly Contribution Needed = `=(B-D)/MAX(1,DATEDIF(TODAY(),C,"M"))` |
| F | % Complete = `=D/B` |
| G | Months Remaining = `=DATEDIF(TODAY(),C,"M")` |
| H | Status = `=IF(D>=B,"✅ Funded",IF(F>=0.75,"🟢 On Track",IF(F>=0.5,"🟡 Behind","🔴 Critical")))` |

**Section B — Monthly Contribution Plan (Rows 22–32)**
- Total Required Monthly = `=SUM(E4:E20)`
- Available for Funds (pull from Dashboard)
- Surplus / Shortfall

**Section C — Progress Visuals**
- For each fund: text-based progress bar
  - `=REPT("■",INT(F4*10))&REPT("□",10-INT(F4*10))&" "&TEXT(F4,"0%")`

**Section D — History Log (Rows 35–80)**
- Columns: Date | Fund Name | Amount Added | Running Balance | Note
- Manual entry rows (user logs contributions here)

---

## SHEET: 📆 Smart Calendar

### Purpose
Visual monthly calendar showing bill due dates, paydays, subscription renewals, and savings transfer reminders.

### Layout

**Section A — Month View Grid (Rows 3–35)**
- Standard 7-column calendar grid (Sun–Sat)
- Year/Month input cell at top (user selects month)
- Formula populates correct day numbers using `DATE()` and `WEEKDAY()`
- Each day cell shows stacked text events from source data below

**Section B — Event Source Table (Rows 40–120)**
This table feeds the calendar grid above. Columns:

| Col | Header |
|-----|---------|
| A | Event Type (dropdown: Bill Due / Payday / Subscription / Savings Transfer / Other) |
| B | Event Name |
| C | Day of Month (1–31) |
| D | Amount ($) |
| E | Recurring? (Yes/No) |
| F | Account / Source |
| G | Color Code (auto by type) |

Pre-populated events (sample):
- Paydays: 1st and 15th (Primary Income), 10th (Freelance sweep)
- Bills: Rent (1st), Electric (5th), Internet (10th), Car payment (15th), Insurance (20th), Credit card (25th)
- Subscriptions: Pull from Subscriptions tab using `IMPORTRANGE` or cross-sheet reference
- Savings Transfers: 1st (Emergency Fund), 15th (Investment)

**Section C — Week Summary**
- Below calendar: totals by week for bills due

**Section D — Upcoming (Next 14 days)**
- Sorted list of next 14 days' events using `FILTER()` and `SORT()`

---

## SHEET: 🔁 Subscriptions

### Purpose
Track all recurring subscriptions with renewal alerts and spend totals.

### Layout

**Section A — Subscription Table (Rows 3–50)**

Columns:
| Col | Header |
|-----|---------|
| A | Service Name |
| B | Category (dropdown: Streaming, Software, Insurance, Finance, Health, Other) |
| C | Billing Cycle (dropdown: Monthly, Annual, Quarterly) |
| D | Amount per Cycle ($) — blue input |
| E | Monthly Cost (normalized) = `=IF(C="Annual",D/12,IF(C="Quarterly",D/3,D))` |
| F | Annual Cost = `=E*12` |
| G | Next Renewal Date — blue input |
| H | Days Until Renewal = `=G-TODAY()` |
| I | Renewal Alert = `=IF(H<=30,"⚠️ Due Soon",IF(H<=7,"🔴 URGENT","✅ OK"))` |
| J | Auto-Renews? (Yes/No dropdown) |
| K | Active? (Yes/No dropdown) |
| L | Account/Card Used |

Pre-populate with 15 sample subscriptions across all categories:
- Streaming: Netflix, Spotify, Disney+, YouTube Premium
- Software: Adobe CC, Microsoft 365, Notion, LastPass
- Insurance: Life, Renters/Home
- Finance: Credit monitoring, investment platform
- Health: Gym, meditation app
- Other: Amazon Prime, iCloud

**Section B — Summary Metrics (Rows 52–62)**
- Total Monthly Spend (active only): `=SUMIF(K4:K50,"Yes",E4:E50)`
- Total Annual Spend: `=SUMIF(K4:K50,"Yes",F4:F50)`
- By Category breakdown (SUMIF per category)
- Count of subscriptions renewing within 30 days
- Upcoming renewals this month (list via FILTER)

**Section C — Category Grouping View (Rows 64–90)**
- Group subs by category with subtotals
- Monthly and annual cost per group

**Section D — Renewal Alert List (Rows 92–110)**
- Auto-sorted list: `=SORT(FILTER(A4:L50, H4:H50<=30), 8, TRUE)`
- Shows only subs renewing within 30 days, sorted by urgency

---

## SHEET: 📈 Net Worth

### Purpose
Track total assets vs liabilities with net worth calculation and history.

### Layout

**Section A — Assets (Rows 3–30)**
Categories:
- Cash & Checking
- Savings Accounts
- Emergency Fund (link from Sinking Funds tab)
- Investment Accounts (401k, IRA, Brokerage)
- Real Estate (market value)
- Vehicle Value(s)
- Other Assets

Columns: Asset Name | Institution | Current Value ($) | Last Updated

**Total Assets** = `=SUM(C4:C30)`

**Section B — Liabilities (Rows 32–55)**
Categories:
- Credit Card Balances (link from Debt Payoff tab)
- Auto Loan(s)
- Mortgage
- Student Loans
- Personal Loans
- Medical Debt
- Other Liabilities

Columns: Liability Name | Institution | Balance Owed ($) | Interest Rate

**Total Liabilities** = `=SUM(C33:C55)`

**Section C — Net Worth Summary (Rows 57–65)**
- Total Assets
- Total Liabilities
- **NET WORTH** = Total Assets − Total Liabilities (large bold cell)
- Debt-to-Asset Ratio = `=TotalLiabilities/TotalAssets`
- Investment Rate = `=InvestmentAssets/TotalAssets`

**Section D — Monthly Snapshot Log (Rows 67–90)**
- Manual log table: Date | Total Assets | Total Liabilities | Net Worth | Notes
- User logs monthly snapshots here for tracking over time

---

## SHEET: 📊 Annual Summary

### Purpose
Year-over-year financial overview pulling from all 12 monthly tabs.

### Layout

**Section A — Monthly Income vs Expenses Table (Rows 3–20)**
- Rows: Jan–Dec + Annual Total
- Columns: Month | Total Income | Total Needs | Total Wants | Total Savings | Net Cash Flow | Savings Rate %
- All values pulled via `INDIRECT()` referencing each monthly tab

**Section B — Annual Totals & Averages (Rows 22–30)**
- Annual Total Income, Expenses, Savings
- Monthly Average Income, Expenses, Savings
- Best Month (income), Worst Month (net cash flow)
- Total Savings Rate for the Year

**Section C — Category Drill-Down (Rows 32–50)**
- Top 5 spending categories across the year
- Year-over-year comparison cells (user inputs prior year actuals in blue)

**Section D — Investment & Savings Rate Tracker (Rows 52–65)**
- Columns: Month | Savings $ | Investment Contribution $ | Savings Rate % | Investment Rate %
- Pulled from monthly Savings section
- Annual totals and rate

**Section E — Month-to-Month Comparison Chart Data (Rows 67–80)**
- Structured table for charting (Income vs Expenses vs Savings per month)
- Note: "Insert chart using this data range: [range]" comment in cell A67

---

## FORMATTING STANDARDS

### Color Palette (Clean Minimal: White + Gray + Blue)
| Element | Color | Hex |
|---------|-------|-----|
| Header backgrounds | Dark navy blue | #1E3A5F |
| Header text | White | #FFFFFF |
| Section headers | Medium blue | #2563EB |
| Section header text | White | #FFFFFF |
| Sub-headers | Light blue | #DBEAFE |
| Sub-header text | Dark blue | #1E3A5F |
| User input cells | White with blue text | Text: #1D4ED8 |
| Input cell background | Light yellow highlight | #FEFCE8 |
| Formula cells | White with black text | #000000 |
| Cross-sheet links | White with green text | #15803D |
| Alternating data rows | White / Light gray | #F9FAFB |
| Positive variance | Light green background | #D1FAE5 |
| Negative variance | Light red background | #FEE2E2 |
| On-track status | Blue | #BFDBFE |
| Section borders | Medium gray | #D1D5DB |
| Total rows | Light gray background | #F3F4F6, bold text |

### Typography & Cell Formatting
- Font: Arial throughout
- Title cells: 18pt bold
- Section headers: 12pt bold
- Data cells: 11pt regular
- Column widths: auto-fit to content + 8px padding
- Row heights: 20px standard, 30px for section headers
- Number format for currency: `$#,##0.00;($#,##0.00);"-"`
- Percentage format: `0.0%`
- Date format: `MMM D, YYYY`
- Freeze top 3 rows on every sheet (header + column labels)
- Freeze column A on monthly tabs

### Conditional Formatting Rules (apply to all relevant sheets)
1. Progress % cells: color scale blue (0%) → green (100%), red if over 100%
2. Variance cells: green fill if positive, red fill if negative, gray if zero
3. Days Until Renewal: red if ≤7, orange if ≤30, green if >30
4. Status/alert cells: already formula-driven with emoji — no additional CF needed
5. Net Cash Flow: green if positive, red if negative

---

## FORMULA STANDARDS

### Cross-Sheet References
- Use `INDIRECT()` on Dashboard to pull from active month: `=INDIRECT("'"&MonthNames&"'!D95")` where MonthNames is a lookup array
- Use direct sheet references on Annual Summary: `=Jan!D95`
- All cross-sheet formula text: green (#15803D)

### Named Ranges (create these)
| Name | Points To |
|------|-----------|
| `ActiveMonth` | Dashboard!B2 |
| `NeedsTarget` | Dashboard!C6 |
| `WantsTarget` | Dashboard!C7 |
| `SavingsTarget` | Dashboard!C8 |
| `ExtraPayment` | DebtPayoff!C2 |
| `Method` | DebtPayoff!E1 |
| `TotalDebt` | DebtPayoff!B22 |
| `TotalAssets` | NetWorth!C31 |
| `TotalLiabilities` | NetWorth!C57 |

### Error Prevention
- Wrap all division formulas: `=IFERROR(numerator/denominator, 0)`
- Wrap all INDIRECT: `=IFERROR(INDIRECT(...), 0)`
- Wrap DATEDIF: `=IFERROR(DATEDIF(TODAY(),C4,"M"), 0)`
- All SUM ranges should include blank rows to allow user additions

---

## SAMPLE DATA (January Tab)

Populate January with this sample data so the user sees a live working example:

**Income:**
- Paycheck 1 (Primary): Expected $4,500 / Actual $4,500
- Paycheck 2 (Primary): Expected $4,500 / Actual $4,500
- Freelance Project: Expected $1,200 / Actual $1,800
- Side Hustle (Etsy/etc): Expected $300 / Actual $425
- Total: Expected $10,500 / Actual $11,225

**Needs (~50%):** Rent $1,800, Electric $120, Internet $65, Mobile $85, Car Payment $485, Car Insurance $175, Fuel $200, Groceries $600, Health Insurance $320, Total ~$3,850

**Wants (~30%):** Dining Out $280, Streaming $65, Shopping $320, Gym $50, Entertainment $150, Amazon $95, Total ~$960

**Savings/Debt (~20%):** Emergency Fund $500, Car Fund $200, Vacation Fund $300, Extra CC Payment $400, 401k contribution $450, Total ~$1,850

**8 Sample Debts (Debt Payoff tab):**
- Chase Visa: $3,200 @ 24.99% APR, Min $65
- Capital One: $1,800 @ 22.49%, Min $45
- Citi Card: $5,500 @ 19.99%, Min $110
- Discover: $900 @ 26.99%, Min $25
- Auto Loan 1: $12,400 @ 7.49%, Min $285
- Auto Loan 2: $8,900 @ 9.99%, Min $195
- Personal Loan: $4,200 @ 14.99%, Min $98
- Medical Bill: $750 @ 0%, Min $50
- Extra Payment: $200/month

---

## DELIVERY

1. Create the Google Sheet using the Workspace MCP/CLI tools
2. Name the file: `💰 Personal Finance Command Center 2025`
3. Share link to the completed Google Sheet
4. Export as `.xlsx` using openpyxl
5. Run `scripts/recalc.py` on the `.xlsx` and confirm zero formula errors
6. Deliver the `.xlsx` download link

After delivery, provide a brief summary of:
- Sheet count and tab names
- How to use the ActiveMonth selector on the Dashboard
- How to toggle Snowball vs Avalanche on the Debt Payoff tab
- Any cells the user should fill in first (in priority order)