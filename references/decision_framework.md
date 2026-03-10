# Decision Framework: When to Ask vs When to Proceed

## Core Principle

Autonomy comes from correctly judging risk, not from asking fewer questions. Ask for HIGH-RISK decisions (where error is expensive). Proceed with SAFE DEFAULTS for everything else. Always DISCLOSE ASSUMPTIONS in your output.

---

## Decision Matrix

| Ambiguity | Risk | Decision | Rationale |
|-----------|------|----------|-----------|
| What the spreadsheet is for | CRITICAL | ASK | Wrong model type = total rebuild |
| Data range/scope | HIGH | ASK | Wrong range = corrupted output |
| Key formulas/calculations | HIGH | ASK | Wrong formula = wrong results |
| Delete/overwrite operations | CRITICAL | ASK ALWAYS | Irreversible data loss |
| Platform (Excel vs Sheets) | MEDIUM | ASK if ambiguous | Default: Excel .xlsx |
| Precision requirements | MEDIUM | ASK for finance | Default: 2 decimal places |
| Industry standards (GAAP etc.) | MEDIUM | ASK if financial | Default: no specific standard |
| Column data types | MEDIUM | PROCEED + INFER | Infer from headers/data; state assumption |
| Sort direction | LOW | PROCEED | Default: ascending |
| Chart type | LOW | PROCEED + SUGGEST | Default: bar chart; user can change |
| Font choice | LOW | PROCEED | Default: Arial |
| Color scheme | LOW | PROCEED | Default: industry-standard (blue inputs, black formulas) |
| Number format details | LOW | PROCEED | Default: currency/percentage as appropriate |
| Sheet naming | LOW | PROCEED | Default: descriptive names (Data, Summary, Dashboard) |

---

## Question Prioritization

### Stage 1: INTENT (Must know — blocks everything)
- What is the user trying to accomplish?
- Is this a financial model, engineering calc, dashboard, tracker, or analysis?
- This determines which domain reference to read

### Stage 2: SCOPE (Must know — blocks execution)
- Which data? How much data?
- How many sheets? What columns?
- Input data format (manual, CSV, existing file)

### Stage 3: CALCULATIONS (Must know — blocks formulas)
- What are the key metrics/KPIs?
- Any specific formulas required?
- Sensitivity analysis or scenario modeling needed?

### Stage 4: PREFERENCES (Nice to have — use defaults)
- Chart types, color preferences, layout details
- Print settings, page orientation
- These are NOT blockers — proceed with defaults

---

## Progressive Disclosure Pattern

```
USER REQUEST
    ↓
Parse intent → Can you understand what they want?
    ├── YES → Move to scope
    └── NO → Ask: "What are you trying to accomplish?"
                  (ONE question only)
    ↓
Check scope → Is the data/structure clear?
    ├── YES → Move to calculations
    └── NO → Ask: "What data should I work with?"
                  (ONE question only)
    ↓
Check calculations → Are the key formulas clear?
    ├── YES → Present plan, proceed
    └── NO → Ask: "What calculations are most important?"
                  (ONE question only)
    ↓
PRESENT PLAN with stated assumptions → Wait for approval
    ↓
BUILD + VERIFY + DELIVER
```

**Never ask more than 2-3 questions before presenting a plan.** The plan itself is a form of question — the user reviews and corrects.

---

## Safe Defaults (Use Without Asking)

### Data Handling
- Blank numeric cells → 0
- Blank text cells → "" (empty string)
- Leading/trailing whitespace → TRIM always
- Duplicate rows → keep first occurrence
- Date format → auto-detect from data

### Formatting
- Font: Arial, 10pt body, 11pt bold headers
- Colors: Blue=inputs, Black=formulas, Green=links
- Currency: `$#,##0;($#,##0);"-"`
- Percentage: `0.0%`
- Negative numbers: parentheses
- Years: text format (no comma)
- Column width: auto-fit to content

### Structure
- Assumptions in dedicated section/sheet (top of model or separate sheet)
- Headers in row 1, data starts row 2
- Totals at bottom of data range
- Summary/output sheets after detail sheets
- Instruction/TOC sheet if 4+ sheets

### Verification
- Tolerance: 0.01 for currency, 1e-6 for engineering
- Always verify SUM, AVERAGE, and key calculated metrics
- Always check cross-sheet references
- Always run recalc.py

---

## When to State Assumptions

After proceeding with defaults, always tell the user what you assumed:

```
"I built this with these assumptions:
- Currency format: USD ($#,##0)
- Fiscal year: calendar year
- Growth rates in the Assumptions sheet
Let me know if you'd like to change any of these."
```

This is cheaper than asking and lets the user correct only what matters.
