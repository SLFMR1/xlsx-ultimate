---
name: xlsx-ultimate
description: >
  Build, edit, and verify Excel spreadsheets with Python-backed backtesting. Every formula is independently recalculated and compared. Use for financial models, engineering calcs, dashboards, budgets, payroll, and any tabular data work. Trigger: Excel, spreadsheet, .xlsx, .csv, budget, financial model, pivot table, formula, or calculation. Do NOT use when the deliverable is a Word doc, PowerPoint, or Python script.
---

# Autonomous Spreadsheet Agent

You plan, build, and verify spreadsheets. Every calculation is independently backtested in Python. You deliver spreadsheets with zero errors.

Spreadsheet errors cost billions (JPMorgan's $6B London Whale, TransAlta's $24M, Fidelity's $2.6B). Your verification pipeline makes errors impossible.

## 5-Phase Pipeline

```
GATHER → PLAN → BUILD → VERIFY → DELIVER
  ↑                         |
  └─── fix & re-verify ─────┘
```

---

## Phase 1: GATHER REQUIREMENTS

Before writing any code, get a complete picture. Use the decision framework in `references/decision_framework.md`.

### Mode Detection

First, determine the mode:
- **Create from scratch**: User wants a new spreadsheet → proceed to questions below
- **Modify existing file**: User provides an .xlsx file to edit → read it first, understand structure, then ask what to change. Always backup: `shutil.copy("input.xlsx", "input_backup.xlsx")`

### Always Ask (blocks execution)

1. **Intent**: What is this spreadsheet for?
2. **Scope**: How much data? How many sheets? What columns?
3. **Key calculations**: What formulas/metrics matter?

### Ask If Ambiguous

4. Platform preference (Excel vs Google Sheets — default: Excel .xlsx)
5. Precision requirements (to the cent, to thousands)
6. Industry standards to follow (GAAP, IFRS, ISO, NEC, ASME)
7. Charts/visualizations needed

### Edge Case Detection

Before building, scan requirements for:
- **Division scenarios** → wrap ALL division formulas in `IFERROR()`
- **Percentage inputs** → validate 0-100% or 0-1 (clarify with user)
- **Date inputs** → validate format, handle Feb 29 / year boundaries
- **Currency inputs** → clarify symbol, decimal places
- **Contradictory totals** → if user gives parts AND total, verify parts sum to total; if not, ASK
- **Circular dependencies** → identify before building (see Circular Reference Handling below)

### Circular Reference Handling

When requirements imply circular logic (debt → interest → NOI → DSCR → debt):
1. **Restructure to avoid** (preferred): Break the loop with an assumption (fix debt amount, calculate interest)
2. **Manual iteration columns**: 3-5 iteration columns that converge — NO actual Excel circular reference
3. **Excel iterative calculation**: Only as last resort. Flag explicitly, add instruction for enabling it in Excel settings.

### Proceed With Safe Defaults

- Font: Arial 10pt (body), 11pt bold (headers)
- Color coding: Blue=inputs, Black=formulas, Green=cross-sheet links
- Number format: `$#,##0;($#,##0);"-"` for currency, `0.0%` for percentages
- Negative numbers in parentheses
- Years as text format (no comma formatting)
- Data validation on input cells

### Domain Detection

```
money/revenue/cost/profit    → Read references/business_domains.md
stress/load/tolerance/units  → Read references/engineering_domains.md
KPI/dashboard/tracking       → Read references/business_domains.md (Dashboards)
data/clean/transform         → Data Analysis (pandas-first approach)
schedule/timeline            → Project Management
inventory/demand/supply      → Supply Chain
salary/headcount             → HR/Workforce
rent/property/cap rate       → Real Estate
```

### Progressive Disclosure

Ask ONE question at a time. Don't overwhelm. State assumptions explicitly: "I'll use [default] — let me know if you'd prefer something different."

---

## Phase 2: PLAN

Create an explicit plan BEFORE writing code. Show it to the user and wait for approval.

### Input Validation (before approving plan)

- Do all parts sum to stated totals? (revenue channels = total revenue)
- Are growth rates plausible? (>100% annual growth → confirm with user)
- Are cost ratios within industry norms? (food cost >50% → flag as unusual)
- If Google Sheets requested → plan ONLY compatible functions (see GSheets table below)

### Plan Template

```
SPREADSHEET PLAN
================
Purpose: [one sentence]
Platform: Excel (.xlsx) / Google Sheets
Mode: Create from scratch / Modify existing
Estimated complexity: Simple / Medium / Complex

SHEET ARCHITECTURE:
  Sheet 1: "[Name]" — [purpose] (columns: [...])
  Sheet 2: "[Name]" — [purpose] (columns: [...])
  Relationships: Sheet2 pulls from Sheet1 via [formulas]

KEY FORMULAS:
  1. [Metric] = [formula logic] (cell range: [location])
  2. [Metric] = [formula logic] (cell range: [location])

ASSUMPTIONS (in dedicated cells, blue font):
  - [Assumption 1]: [value] (source: [reference])

DATA VALIDATION:
  - [Cell range]: [rule] (e.g., dropdown, number range)

CHARTS (if applicable):
  - [Chart type]: [data source] → placed on [sheet]

VERIFICATION STRATEGY:
  - [Check 1]: Python will independently calculate [what]
  - [Check 2]: [Domain-specific check, e.g., A=L+E for balance sheets]
  Tolerance: [abs_tol for currency, rel_tol for engineering]
```

Wait for user approval. Adjust if needed.

---

## Phase 3: BUILD

### Tech Stack

- **openpyxl**: Create/edit .xlsx with formulas, formatting, charts
- **pandas**: Data manipulation, bulk operations, analysis
- **LibreOffice**: Formula recalculation via `scripts/recalc.py`
- For code patterns, see `references/build_patterns.md`

### Non-Negotiable Rules

**1. Excel formulas, never hardcoded values:**
```python
# WRONG:  sheet['B10'] = 7350          # hardcoded number
# WRONG:  sheet['B10'] = sum(values)   # Python calculation written as value
# RIGHT:  sheet['B10'] = '=SUM(B2:B9)' # Excel formula string
# RIGHT:  sheet['D4'] = '=PI()*(B3/2)^2' # Excel formula for area
```
All calculations must be Excel formula STRINGS (starting with "=") so the spreadsheet stays dynamic. Every calculated cell must contain a formula string, never a Python-computed number.

**2. Absolute references where needed:**
```python
# WRONG:  '=B2*E1'     (E1 shifts when copied)
# RIGHT:  '=B2*$E$1'   (E1 stays fixed)
```

**3. Assumptions separated from calculations:**
All inputs in dedicated cells (blue font). Formulas reference those cells. No magic numbers.

**4. Defensive formulas (MANDATORY — every single division formula MUST be wrapped):**
```python
# EVERY formula containing "/" MUST be wrapped in IFERROR. No exceptions.
'=IFERROR(B5/B6, 0)'
'=IFERROR(Revenue/Employees, 0)'
'=IFERROR(D2*0.25, 0)'  # Even if divisor seems safe, wrap it

# After building ALL formulas, do a self-check:
# Search your code for "/" in any formula string. If it's not inside IFERROR(), fix it.

# Handle empty cells:
'=IF(ISBLANK(B5), 0, B5*$C$1)'
# Negative protection for non-negative metrics:
'=MAX(0, B5-C5)'
```
**IFERROR audit rule**: Before moving to Phase 4, grep your own build script for any formula containing `/` that is NOT wrapped in `IFERROR()`. Fix every one.

**5. Industry-standard color coding:**
Blue=inputs, Black=formulas, Green=cross-sheet links, Yellow bg=key assumptions.

**6. Proper number formatting:**
Currency: `$#,##0;($#,##0);"-"` | Percent: `0.0%` | Multiples: `0.0"x"` | Years: text `@`

### Build Workflow

1. Create workbook structure (sheets, headers, column widths)
2. Populate static data and assumptions
3. Write formulas (cell references, not hardcoded)
4. Apply formatting (fonts, colors, borders, number formats)
5. Add data validation (dropdowns, number ranges, date constraints — see `references/build_patterns.md`)
6. Add conditional formatting (if applicable):
   ```python
   from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, IconSetRule
   # Heat map (green-yellow-red):
   ws.conditional_formatting.add('B2:B100', ColorScaleRule(
       start_type='min', start_color='63BE7B',
       mid_type='percentile', mid_value=50, mid_color='FFEB84',
       end_type='max', end_color='F8696B'))
   # Highlight overdue (formula-based CF with absolute row, relative column):
   ws.conditional_formatting.add('A2:H100', CellIsRule(
       operator='lessThan', formula=['TODAY()'], fill=PatternFill(bgColor='FFC7CE')))
   # Data bars:
   ws.conditional_formatting.add('C2:C100', DataBarRule(start_type='min', end_type='max', color='638EC6'))
   ```
7. Create charts (if requested or if data has clear visual dimension):
   **IMPORTANT: "Charts" means openpyxl chart OBJECTS (BarChart, LineChart, etc.), NOT just conditional formatting. If the task says "chart" or "Gantt chart", you MUST create an actual chart object with `ws.add_chart()`. CF-based visualizations are supplementary, not a replacement for chart objects.**
   ```python
   from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, Reference
   # ALWAYS use absolute references for chart data ranges
   # ALWAYS set title, axis labels, legend
   # For dynamic data, use named ranges so user can add rows:
   #   wb.defined_names.new("SalesData", attr_text="'Data'!$A$1:$D$1000")
   # Position chart: ws.add_chart(chart, "F2") — never overlap data
   # Common types: BarChart, LineChart, PieChart (max 8 segments), ScatterChart
   # Combo: BarChart + LineChart on secondary axis (c2.y_axis = chart.y_axis)
   # For Gantt/timeline: Use a stacked BarChart (invisible start + colored duration)
   ```
8. Lock formula cells, unlock input cells:
   ```python
   from openpyxl.styles import Protection
   for row in ws.iter_rows():
       for cell in row:
           if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
               cell.protection = Protection(locked=True)
           elif cell.value is not None:
               cell.protection = Protection(locked=False)
   ws.protection.sheet = True
   ws.protection.enable()
   ```
   Skip protection if user explicitly requests fully editable template.
9. Print layout (for sheets >50 rows):
   ```python
   ws.freeze_panes = 'A2'
   ws.print_title_rows = '1:1'  # repeat header on each page
   ws.sheet_properties.pageSetUpPr = PrintPageSetup(fitToPage=True)
   ws.oddFooter.center.text = 'Page &P of &N'
   # Landscape for >6 columns:
   ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
   ```
10. Save file
11. **Run recalc** (if available):
    ```bash
    cd scripts && python recalc.py ../output.xlsx 30 && cd ..
    ```
    If recalc fails or LibreOffice unavailable, log warning and proceed to Phase 4. Shadow calculations are the primary verification; recalc is supplementary.
12. If `errors_found` → fix errors → recalc again → repeat until `success`
13. Proceed to Phase 4

### Modifying Existing Files

When modifying an existing .xlsx (not creating from scratch):
1. `shutil.copy("input.xlsx", "input_backup.xlsx")` — always backup
2. `wb = load_workbook("input.xlsx")` — load existing
3. Read existing structure: sheets, columns, formulas, formatting
4. Only modify what user specifically asked to change
5. Preserve existing formatting, formulas, and data validation
6. Phase 4: verify original formulas still work + new formulas correct

---

## Phase 4: VERIFY (Python Backtesting)

This is the entire point of this skill. Without real verification, the spreadsheet is just guesswork. Structural checks (formula exists, no #REF errors) are NOT verification. Real verification means: **recalculate every key value independently in Python and compare it to the Excel output.**

### The Verification Script You Must Write

For EVERY spreadsheet, write and run a **custom Python verification script**. The script follows this exact pattern:

```python
import json, math, re
from openpyxl import load_workbook

# === STEP 1: Formula count check ===
wb_formulas = load_workbook("output.xlsx", data_only=False)
formula_count = 0
for sn in wb_formulas.sheetnames:
    for row in wb_formulas[sn].iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                formula_count += 1
assert formula_count >= 10, f"FAIL: Only {formula_count} formulas. Rebuild with real Excel formulas."

# === STEP 2: Cross-sheet reference validation ===
sheet_ref_pattern = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_]+))!")
for sn in wb_formulas.sheetnames:
    for row in wb_formulas[sn].iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith("="):
                for match in sheet_ref_pattern.finditer(cell.value):
                    ref_sheet = match.group(1) or match.group(2)
                    assert ref_sheet in wb_formulas.sheetnames, \
                        f"Broken ref: {sn}!{cell.coordinate} → '{ref_sheet}' not found!"
wb_formulas.close()

# === STEP 3: Load calculated values ===
wb = load_workbook("output.xlsx", data_only=True)

# === STEP 4: Read INPUT values from spreadsheet ===
assumptions = wb["Assumptions"]
price = assumptions["B5"].value
# ... read ALL input parameters

# === STEP 5: INDEPENDENTLY recalculate in Python ===
python_mrr = customers * price
# ... recalculate EVERY key metric

# === STEP 6: Read EXCEL values and COMPARE ===
checks = []
def compare(name, python_val, excel_val, abs_tol=0.01, rel_tol=1e-6):
    if python_val is None or excel_val is None:
        return {"name": name, "python": python_val, "excel": excel_val, "match": False, "diff": None}
    match = math.isclose(python_val, excel_val, abs_tol=abs_tol, rel_tol=rel_tol)
    return {"name": name, "python": round(python_val, 6), "excel": round(excel_val, 6),
            "match": match, "diff": round(abs(python_val - excel_val), 6)}

checks.append(compare("MRR Month 1", python_mrr, excel_mrr, abs_tol=0.01))
# ... at least 5 comparisons per sheet

# === STEP 7: Sanity bounds ===
# Catch both Python bugs AND Excel formula errors
assert python_revenue >= 0, "Revenue should be non-negative"
# assert 0 <= python_margin <= 1, "Margin between 0-100%"
# assert python_fos >= 1.0, "Factor of safety must be >= 1.0"

# === STEP 8: Conditional formatting audit ===
cf_audit = []
for sn in wb.sheetnames:
    ws = wb[sn]
    for cf in ws.conditional_formatting:
        cf_audit.append({"sheet": sn, "range": str(cf.sqref), "rule_count": len(cf.rules)})

# === STEP 9: Chart audit ===
chart_audit = []
for sn in wb.sheetnames:
    ws = wb[sn]
    for chart in ws._charts:
        chart_audit.append({"sheet": sn, "type": chart.__class__.__name__, "title": str(chart.title)})

# === STEP 10: SAVE verification_report.json (NON-NEGOTIABLE) ===
report = {
    "file": "output.xlsx",
    "formula_count": formula_count,
    "checks": checks,
    "summary": {
        "passed": sum(1 for c in checks if c["match"]),
        "total": len(checks),
        "pass_rate": sum(1 for c in checks if c["match"]) / max(len(checks), 1)
    },
    "confidence": "HIGH" if all(c["match"] for c in checks) else "LOW",
    "recalc_status": "success",  # or "skipped — LibreOffice not available"
    "conditional_formatting": cf_audit,
    "charts": chart_audit
}
with open("verification_report.json", "w") as f:
    json.dump(report, f, indent=2)
print(f"Saved verification_report.json: {report['summary']}")
wb.close()
```

**This entire script is non-negotiable.** The JSON save at the end MUST happen. Do NOT only print to stdout.

**CRITICAL: verification_report.json schema requirements:**
- The key MUST be `"checks"` (not `"structural_checks"`, not `"validations"`, not `"results"`)
- Each check MUST have: `{"name": str, "python": number, "excel": number, "match": bool, "diff": number}`
- The `"summary"` MUST have: `{"passed": int, "total": int, "pass_rate": float}`
- Structural-only checks (e.g., "formula count >= 10") do NOT count as real verification — they are supplementary
- You MUST have at least 5 checks where `"python"` and `"excel"` are actual numerical values compared via `math.isclose()`

### What Counts as Real Verification

**YES (real verification):**
- "Python calculated MRR = 8109.50, Excel cell E4 = 8109.50 → MATCH"
- "Python gasket area = π*(275/2)² = 59396 mm², Excel D4 = 49087 mm² → MISMATCH"
- "Python weighted pipeline = $3,150,000, Excel G15 = $3,150,000 → MATCH"

**NO (not verification, just structural checks):**
- "Cell E4 contains a formula starting with =" ← proves nothing
- "Found keyword MRR in spreadsheet" ← proves nothing
- "51 formula checks passed" ← meaningless if none compare values

### Minimum Verification Requirements

1. **Formula count check**: ≥10 formulas or STOP and rebuild
2. **Cross-sheet validation**: All sheet references resolve to existing sheets
3. **Read actual Excel cell values** (data_only=True)
4. **Independent Python calculation** from the same inputs
5. **Numerical comparison** with `math.isclose()`
6. **At least 5 key value comparisons** per sheet (for large datasets: verify first, middle, last row + aggregates)
7. **Sanity bounds**: Revenue ≥0, margins 0-100%, FoS ≥1, etc.
8. **Save verification_report.json** to same directory as xlsx

### Tolerance Levels

- Currency: `abs_tol=0.01` (to the cent)
- Percentages: `abs_tol=0.0001`
- Engineering: `rel_tol=1e-6`
- Exact counts: `abs_tol=0` (exact match)

### Domain-Specific Checks

**Financial models:** MRR/ARR math, customer churn progression, cash balance = opening + revenue - costs, balance sheet A=L+E.

**Engineering:** Recalculate forces/stresses from input dimensions and material properties. Check units. Verify factor of safety.

**Dashboards:** SUMIF/COUNTIF totals match manual count. Weighted values = sum(value × probability).

**Accounting:** Debits = Credits on trial balance. A = L + E always.

If any mismatch → proceed to Error Recovery. Do NOT deliver.

---

## Phase 4b: ERROR RECOVERY

When verification fails:

### Auto-Fix (No User Input Needed)

- **Off-by-one range**: Expand `=SUM(B2:B9)` to `=SUM(B2:B10)` if row 10 contains data
- **Missing absolute ref**: Add `$` where formula should be fixed
- **Division by zero**: Wrap with `=IFERROR(formula, 0)`
- **Empty formula result**: Check source cells for data, fix references
- After auto-fix → re-run FULL verification (including verification_report.json save)

### Ask User (Ambiguous Intent)

- "I calculated X=5,000 but expected X=5,050. Should the formula include [specific row/column]?"
- "The discount rate cell is empty. What value should I use?"
- "This creates a circular reference. Should I restructure the calculation?"

### Iteration Rule

**Never deliver a spreadsheet with failed verification.** Fix → re-verify → repeat. Maximum 3 fix-and-verify cycles. After 3 failures on the same check, document and ask user.

---

## Phase 5: DELIVER

### Delivery Checklist

- [ ] File saved to output directory
- [ ] **verification_report.json saved to SAME directory as xlsx** (MANDATORY)
- [ ] Python verification: 100% pass (or documented exceptions)
- [ ] Formatting complete (colors, fonts, number formats, borders)
- [ ] Data validation in place
- [ ] Conditional formatting applied (if applicable)
- [ ] Charts render correctly (if applicable)
- [ ] Cell protection set (formula cells locked, input cells unlocked)
- [ ] Print layout set for sheets >50 rows (freeze panes, print titles)
- [ ] Assumptions documented

**If verification_report.json is missing, DO NOT deliver. Go back to Phase 4.**

### Deliver Format

```
[View your spreadsheet](computer:///path/to/filename.xlsx)

Verification: [N]/[N] checks passed. All calculations independently verified.
[Brief summary of what was built]
```

### Iteration

If user wants changes: understand feedback → update plan → modify → re-verify → deliver.
Never skip verification after changes, even for "small" edits.

---

## Google Sheets Compatibility

If user requests Google Sheets, use ONLY compatible functions:

```
Excel-Only Function  → Google Sheets Replacement
XLOOKUP              → INDEX(MATCH())
LET                  → Inline the expression
LAMBDA               → Use helper cells
IFS                  → Nested IF()
SWITCH               → Nested IF() or VLOOKUP on helper table
XMATCH               → MATCH()
VSTACK / HSTACK      → Manual ranges
Dynamic arrays       → ARRAYFORMULA() wrapper
Structured refs      → Use A1 notation
SUBTOTAL mode 101+   → SUMPRODUCT workaround (ignore hidden rows)
```

Phase 4: If GSheets mode, scan ALL formulas for Excel-only functions. Flag any found.

---

## Reference Files

Read ONLY the files relevant to the current task. Do NOT load all references upfront.

| File | When to Read | Priority |
|------|-------------|----------|
| `references/decision_framework.md` | ALWAYS read first — controls question flow and safe defaults | Phase 1 |
| `references/business_domains.md` | Financial models, FP&A, Accounting, Dashboards, Sales, HR, Supply Chain, Real Estate | Phase 2 (if business) |
| `references/engineering_domains.md` | Mechanical, Electrical, Civil, Chemical, Manufacturing, Quality, Systems Engineering | Phase 2 (if engineering) |
| `references/build_patterns.md` | Code patterns, charts, data validation, pivot tables, sensitivity analysis, Google Sheets compat | Phase 3 |
| `references/formula_reference.md` | Complex formulas, function syntax, Excel vs Google Sheets differences | Phase 3 (as needed) |
| `references/verification_guide.md` | Shadow calculation patterns, tolerance handling, numerical precision | Phase 4 |
| `scripts/verify_spreadsheet.py` | Run generic verification; also contains `npv_python()`, `irr_python()`, `pmt_python()` helpers | Phase 4 |

---

## Code Style

Write minimal, concise Python. No unnecessary comments or print statements. Break complex operations into small testable functions. Use type hints for verification functions.
