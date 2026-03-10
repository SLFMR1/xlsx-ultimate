# Business Domain Reference

## Table of Contents
1. [Financial Modeling](#financial-modeling)
2. [FP&A](#fpa)
3. [Accounting](#accounting)
4. [Dashboards & BI](#dashboards)
5. [Sales & CRM](#sales)
6. [HR & Workforce](#hr)
7. [Supply Chain](#supply-chain)
8. [Real Estate](#real-estate)
9. [Project Management](#project-management)

---

## Financial Modeling {#financial-modeling}

### Model Types

**DCF (Discounted Cash Flow):**
- Project free cash flows 5-10 years
- Calculate terminal value (Gordon Growth or Exit Multiple)
- Discount at WACC
- Key formulas: `NPV()`, `XNPV()`, `IRR()`, `XIRR()`
- Structure: Assumptions → Revenue Build → P&L → Balance Sheet → Cash Flow → DCF → Sensitivity

**LBO (Leveraged Buyout):**
- Purchase price and financing structure
- Debt schedule (tranches, interest, amortization)
- Operating model (revenue, EBITDA projections)
- Exit analysis (multiple scenarios)
- Returns: IRR and MOIC (Multiple on Invested Capital)

**Three-Statement Model:**
- Income Statement → Balance Sheet → Cash Flow Statement
- Linked via: Net Income, D&A, CapEx, Working Capital changes
- Balance sheet MUST balance: Assets = Liabilities + Equity
- Cash flow reconciliation: Opening Cash + Net Change = Closing Cash

**Merger/Accretion-Dilution:**
- Combined pro forma financials
- Purchase price allocation
- Goodwill calculation
- EPS accretion/dilution analysis

### Financial Model Structure (Left-to-Right)

```
Sheet 1: Assumptions (all inputs in blue, growth rates, margins, multiples)
Sheet 2: Revenue Build (detailed revenue drivers)
Sheet 3: Income Statement (Revenue → EBITDA → Net Income)
Sheet 4: Balance Sheet (Assets, Liabilities, Equity)
Sheet 5: Cash Flow (Operating, Investing, Financing)
Sheet 6: Debt Schedule (if applicable)
Sheet 7: DCF / Valuation
Sheet 8: Sensitivity Analysis
Sheet 9: Output / Summary
```

### Key Financial Formulas

```
Revenue Growth: =(Current - Prior) / Prior
Gross Margin: =(Revenue - COGS) / Revenue
EBITDA Margin: =EBITDA / Revenue
Net Income Margin: =Net_Income / Revenue
EV/EBITDA Multiple: =Enterprise_Value / EBITDA
P/E Ratio: =Share_Price / EPS
WACC: =E/V * Re + D/V * Rd * (1-T)
Terminal Value (Gordon): =FCF * (1+g) / (WACC - g)
Working Capital: =Current_Assets - Current_Liabilities
Days Sales Outstanding: =Receivables / Revenue * 365
Days Payable Outstanding: =Payables / COGS * 365
Inventory Turnover: =COGS / Average_Inventory
```

### Sensitivity Analysis

**One-Way:** Change single variable, observe impact on output
- Use Excel Data Tables: Data → What-If Analysis → Data Table
- Common: Revenue growth ±2%, Discount rate ±1%, Exit multiple ±1x

**Two-Way:** Change two variables simultaneously
- Matrix format: Variable 1 across columns, Variable 2 down rows
- Output in intersection cells

**Tornado Chart:** Rank variables by impact magnitude
- Show which assumptions matter most

**Monte Carlo:** (requires Python)
- Define probability distributions for key assumptions
- Run 10,000+ simulations
- Output: probability distribution of outcomes, confidence intervals

### Verification Checks for Financial Models

```python
# 1. Balance sheet balances
assert abs(total_assets - (total_liabilities + total_equity)) < 0.01

# 2. Cash flow reconciles
assert abs(opening_cash + net_cash_flow - closing_cash) < 0.01

# 3. Net income flows correctly
assert income_statement_net_income == cash_flow_starting_net_income

# 4. D&A is consistent
assert pl_depreciation == cf_depreciation == bs_accumulated_depreciation_change

# 5. CapEx is consistent
assert cf_capex == bs_ppe_change + depreciation

# 6. Working capital changes match
assert cf_wc_change == (current_assets_change - current_liabilities_change)

# 7. Debt schedule matches
assert cf_debt_change == bs_debt_change

# 8. Shares outstanding consistent
assert eps * shares == net_income
```

---

## FP&A {#fpa}

### Budget vs Actual Variance Analysis

**Sheet Structure:**
```
Column A: Line Item (Revenue, COGS, SG&A, etc.)
Column B: Budget (Annual plan)
Column C: Actual (Period results)
Column D: Variance ($) =C-B
Column E: Variance (%) =(C-B)/ABS(B)
Column F: Commentary (text explanation)
```

**Key Formulas:**
```
Favorable variance (revenue): =IF(Actual>Budget, "Favorable", "Unfavorable")
Favorable variance (cost): =IF(Actual<Budget, "Favorable", "Unfavorable")
Variance %: =IFERROR((Actual-Budget)/ABS(Budget), 0)
```

### Rolling Forecast

- Always show next 12-18 months regardless of fiscal year
- Replace actuals as periods close
- Use OFFSET or INDEX for dynamic date ranges
- Separate "Locked" (actuals) from "Forecast" columns

### Revenue Modeling

**Top-Down:** TAM → SAM → SOM → Market share → Revenue
**Bottom-Up:** Units × Price, or Customers × ARPU × Retention
**Cohort-Based:** New customers per period, retention curve, LTV

---

## Accounting {#accounting}

### Trial Balance

```
Column A: Account Code
Column B: Account Name
Column C: Debit Balance
Column D: Credit Balance
TOTAL: Sum(Debits) MUST equal Sum(Credits)
```

### Depreciation Methods

```
Straight-Line: =(Cost - Salvage) / Useful_Life
Declining Balance: =Book_Value × Rate
Double-Declining: =Book_Value × (2 / Useful_Life)
Sum-of-Years: =Depreciable_Base × (Remaining_Life / Sum_of_Years)
Units-of-Production: =(Cost - Salvage) × (Units_This_Period / Total_Units)
```

### Amortization Schedule (Loan)

```
Payment: =PMT(Rate/12, Periods, -Principal)
Interest: =Beginning_Balance × Rate/12
Principal: =Payment - Interest
Ending Balance: =Beginning_Balance - Principal_Payment
```

### Reconciliation Template

```
Bank Balance (per statement)
+ Deposits in transit
- Outstanding checks
= Adjusted Bank Balance

Book Balance (per ledger)
+ Interest earned
- Bank charges
± Error corrections
= Adjusted Book Balance

Adjusted Bank Balance MUST equal Adjusted Book Balance
```

---

## Dashboards & BI {#dashboards}

### KPI Dashboard Design

**Layout Principles:**
- 5-7 key metrics maximum per dashboard
- Top row: Summary KPIs (big numbers with trend arrows)
- Middle: Charts (trends, comparisons)
- Bottom: Detail tables
- Consistent color coding: Green=good, Red=bad, Yellow=warning

**Common KPIs by Function:**
- Sales: Revenue, Pipeline value, Win rate, Avg deal size, Sales cycle length
- Marketing: CAC, LTV, Conversion rate, MQL→SQL ratio, ROAS
- Finance: Gross margin, Operating margin, Cash runway, Burn rate, ARR
- Operations: Utilization rate, SLA compliance, Ticket resolution time
- HR: Turnover rate, Time to hire, eNPS, Headcount vs plan

**Chart Selection:**
- Trend over time → Line chart
- Comparison across categories → Bar/Column chart
- Part of whole → Pie chart (use sparingly, max 5 segments)
- Distribution → Histogram
- Correlation → Scatter plot
- Flow/waterfall → Waterfall chart

### Conditional Formatting Patterns

```python
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule

# Traffic light (Red/Yellow/Green)
ws.conditional_formatting.add('B2:B20',
    CellIsRule(operator='greaterThan', formula=['0.1'],
              fill=PatternFill(bgColor='00FF00')))

# Data bars for visual magnitude
ws.conditional_formatting.add('C2:C20',
    DataBarRule(start_type='min', end_type='max',
               color='638EC6'))

# Color scale (gradient)
ws.conditional_formatting.add('D2:D20',
    ColorScaleRule(start_type='min', start_color='F8696B',
                   mid_type='percentile', mid_value=50, mid_color='FFEB84',
                   end_type='max', end_color='63BE7B'))
```

---

## Sales & CRM {#sales}

### Pipeline Tracker

```
Columns: Deal Name, Account, Owner, Stage, Amount, Close Date, Probability, Weighted Amount
Weighted Amount: =Amount × Probability
Pipeline Total: =SUM(Weighted_Amount)

Stages: Prospect(10%) → Qualified(25%) → Proposal(50%) → Negotiation(75%) → Closed(100%)
```

### Commission Calculator

```
Base commission: =Revenue × Commission_Rate
Tiered: =SUMPRODUCT((Revenue>Tier_Thresholds)*(Tier_Rates)*MIN(Revenue-Tier_Thresholds, Tier_Widths))
Accelerator: =IF(Revenue>Quota, Base + (Revenue-Quota)*Accelerator_Rate, Base)
```

---

## HR & Workforce {#hr}

### Headcount Planning

```
Columns: Department, Role, Level, FTE, Annual Salary, Benefits %, Total Cost, Start Month
Total Cost: =Salary × (1 + Benefits_Rate) × FTE × (Months_Remaining / 12)
Department Total: =SUMIFS(Total_Cost, Department, "Engineering")
```

### Compensation Analysis

```
Compa-Ratio: =Actual_Salary / Midpoint_Salary
Range Penetration: =(Salary - Range_Min) / (Range_Max - Range_Min)
```

---

## Supply Chain {#supply-chain}

### EOQ (Economic Order Quantity)

```
EOQ: =SQRT(2 × Annual_Demand × Order_Cost / Holding_Cost)
Reorder Point: =Daily_Demand × Lead_Time + Safety_Stock
Safety Stock: =Z_Score × STDEV(Daily_Demand) × SQRT(Lead_Time)
```

### Demand Forecasting

```
Simple Moving Average: =AVERAGE(OFFSET(cell, -periods, 0, periods, 1))
Exponential Smoothing: =Alpha × Actual + (1-Alpha) × Previous_Forecast
Linear Trend: =FORECAST(target_x, known_y, known_x)
Seasonal Index: =Period_Average / Grand_Average
```

---

## Real Estate {#real-estate}

### Pro Forma Structure

```
Gross Potential Rent (from Rent Roll)
- Vacancy Loss (typically 5-10%)
= Effective Gross Income
+ Other Income (parking, laundry, fees)
= Total Revenue
- Operating Expenses (insurance, taxes, maintenance, management)
= Net Operating Income (NOI)
- Debt Service (P&I)
= Cash Flow Before Tax
```

### Key Metrics

```
Cap Rate: =NOI / Property_Value
Cash-on-Cash Return: =Annual_Cash_Flow / Total_Cash_Invested
Debt Service Coverage Ratio: =NOI / Annual_Debt_Service
Loan-to-Value: =Loan_Amount / Property_Value
Gross Rent Multiplier: =Property_Price / Gross_Annual_Rent
```

---

## Project Management {#project-management}

### Gantt Chart Data Structure

```
Columns: Task ID, Task Name, Start Date, End Date, Duration, Dependencies, Owner, Status, % Complete
Duration: =End_Date - Start_Date
Visual: Use conditional formatting to create horizontal bars based on date ranges
```

### Resource Allocation

```
Utilization Rate: =Billable_Hours / Available_Hours
Over-allocation: =IF(Assigned_Hours > Available_Hours, "OVER", "OK")
```
