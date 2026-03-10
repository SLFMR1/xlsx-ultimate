# Formula Reference & Platform Guide

## Table of Contents
1. [Essential Excel Functions](#essential)
2. [Financial Functions](#financial)
3. [Lookup & Reference](#lookup)
4. [Statistical Functions](#statistical)
5. [Date & Time Functions](#datetime)
6. [Text Functions](#text)
7. [Logical & Error Handling](#logical)
8. [Array & Dynamic](#array)
9. [Excel vs Google Sheets Differences](#differences)
10. [Named Ranges & Best Practices](#naming)

---

## Essential Excel Functions {#essential}

### Aggregation
```
SUM(range)                    Sum of values
SUMIF(range, criteria, sum_range)   Sum with condition
SUMIFS(sum_range, range1, criteria1, ...)  Sum with multiple conditions
SUMPRODUCT(array1, array2)    Sum of products (very versatile)
SUBTOTAL(function_num, range) Aggregation ignoring hidden/filtered rows
  function_num: 1=AVG, 2=COUNT, 3=COUNTA, 9=SUM, etc.
AGGREGATE(function, options, range)  Advanced aggregation with error handling
```

### Counting
```
COUNT(range)       Count numeric cells
COUNTA(range)      Count non-empty cells
COUNTBLANK(range)  Count empty cells
COUNTIF(range, criteria)    Count with condition
COUNTIFS(range1, criteria1, ...)  Count with multiple conditions
```

---

## Financial Functions {#financial}

### Time Value of Money
```
PV(rate, nper, pmt, [fv], [type])      Present Value
FV(rate, nper, pmt, [pv], [type])      Future Value
PMT(rate, nper, pv, [fv], [type])      Payment
IPMT(rate, per, nper, pv, [fv], [type])  Interest portion
PPMT(rate, per, nper, pv, [fv], [type])  Principal portion
NPER(rate, pmt, pv, [fv], [type])      Number of periods
RATE(nper, pmt, pv, [fv], [type])      Interest rate
```

### Investment Analysis
```
NPV(rate, value1, value2, ...)         Net Present Value
  NOTE: NPV assumes first cash flow is at period 1, not period 0
  Correct usage: =Initial_Investment + NPV(rate, CF1:CFn)

XNPV(rate, values, dates)             NPV with specific dates
IRR(values, [guess])                   Internal Rate of Return
XIRR(values, dates, [guess])           IRR with specific dates
MIRR(values, finance_rate, reinvest_rate)  Modified IRR
```

### Depreciation
```
SLN(cost, salvage, life)               Straight-line
DB(cost, salvage, life, period, [month])  Declining balance
DDB(cost, salvage, life, period, [factor])  Double declining
SYD(cost, salvage, life, period)       Sum-of-years digits
```

---

## Lookup & Reference {#lookup}

### Modern Approach: INDEX/MATCH (Preferred)

```
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))

Two-way lookup:
=INDEX(data_range,
       MATCH(row_lookup, row_headers, 0),
       MATCH(col_lookup, col_headers, 0))

Multi-criteria:
=INDEX(return_range,
       MATCH(1, (criteria1_range=value1)*(criteria2_range=value2), 0))
  NOTE: This is an array formula (Ctrl+Shift+Enter in older Excel)
```

### XLOOKUP (Excel 365+ Only, NOT in Google Sheets)

```
=XLOOKUP(lookup_value, lookup_array, return_array,
         [if_not_found], [match_mode], [search_mode])

match_mode: 0=exact, -1=exact or next smaller, 1=exact or next larger
search_mode: 1=first-to-last, -1=last-to-first, 2=binary ascending
```

### Legacy: VLOOKUP

```
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
  range_lookup: FALSE for exact match (almost always use FALSE)
  Limitation: Can only look right (return column must be to right of lookup column)
```

### Other Reference Functions

```
INDIRECT(ref_text)    Create reference from text string
OFFSET(ref, rows, cols, [height], [width])  Dynamic range (volatile!)
ROW([reference])      Row number of reference
COLUMN([reference])   Column number of reference
ADDRESS(row, col)     Create cell address string
```

---

## Statistical Functions {#statistical}

### Central Tendency
```
AVERAGE(range)         Mean
AVERAGEIF(range, criteria, avg_range)
AVERAGEIFS(avg_range, range1, criteria1, ...)
MEDIAN(range)          Median (50th percentile)
MODE.SNGL(range)       Most frequent value
TRIMMEAN(range, percent)  Mean excluding outliers
```

### Dispersion
```
STDEV.S(range)         Sample standard deviation (n-1)
STDEV.P(range)         Population standard deviation (n)
VAR.S(range)           Sample variance
VAR.P(range)           Population variance
```

### Distribution & Probability
```
NORM.DIST(x, mean, stdev, cumulative)      Normal distribution
NORM.INV(probability, mean, stdev)          Inverse normal
NORM.S.DIST(z, cumulative)                  Standard normal
T.DIST(x, deg_freedom, cumulative)          T-distribution
T.INV(probability, deg_freedom)              Inverse T
CONFIDENCE.NORM(alpha, stdev, size)          Confidence interval
```

### Regression & Correlation
```
SLOPE(known_y, known_x)           Regression slope
INTERCEPT(known_y, known_x)       Regression intercept
RSQ(known_y, known_x)             R-squared
CORREL(array1, array2)             Correlation coefficient
FORECAST(x, known_y, known_x)     Predicted value
FORECAST.LINEAR(x, known_y, known_x)  Same, explicit linear
TREND(known_y, known_x, new_x)    Array of predicted values
LINEST(known_y, known_x, const, stats)  Full regression statistics
```

### Percentiles & Ranking
```
PERCENTILE.INC(array, k)     k-th percentile (inclusive)
QUARTILE.INC(array, quart)   Quartile (0=min, 1=Q1, 2=median, 3=Q3, 4=max)
RANK.EQ(number, ref, [order])   Rank (ties get same rank)
PERCENTRANK.INC(array, x)    Percentile rank of value
LARGE(array, k)              k-th largest value
SMALL(array, k)              k-th smallest value
```

---

## Date & Time Functions {#datetime}

```
TODAY()                     Current date
NOW()                       Current date and time
DATE(year, month, day)      Create date
YEAR(date), MONTH(date), DAY(date)  Extract components
EDATE(start_date, months)   Date + months
EOMONTH(start_date, months) End of month + months
NETWORKDAYS(start, end, [holidays])  Working days between dates
DATEDIF(start, end, "unit") Difference in unit (Y, M, D, YM, YD, MD)
  NOTE: DATEDIF is undocumented in Excel but works. Not available in all Sheets versions.
WEEKDAY(date, [return_type])  Day of week (1=Sun default)
```

---

## Text Functions {#text}

```
CONCATENATE(text1, text2, ...)  or  text1 & text2   Join text
TEXTJOIN(delimiter, ignore_empty, text1, ...)  Join with delimiter
LEFT(text, num_chars)       First N characters
RIGHT(text, num_chars)      Last N characters
MID(text, start_num, num_chars)  Extract from middle
LEN(text)                   Length
TRIM(text)                  Remove extra spaces
CLEAN(text)                 Remove non-printable chars
UPPER(text), LOWER(text), PROPER(text)   Case conversion
FIND(find_text, within_text, [start])    Case-sensitive position
SEARCH(find_text, within_text, [start])  Case-insensitive position
SUBSTITUTE(text, old, new, [instance])   Replace text
TEXT(value, format)         Format number as text
VALUE(text)                 Convert text to number
```

---

## Logical & Error Handling {#logical}

```
IF(condition, true_value, false_value)
IFS(condition1, value1, condition2, value2, ...)   Multiple conditions (365+)
AND(condition1, condition2, ...)
OR(condition1, condition2, ...)
NOT(condition)
SWITCH(expression, value1, result1, ..., [default])

IFERROR(value, value_if_error)     Catch any error
IFNA(value, value_if_na)           Catch only #N/A
ISERROR(value)                     TRUE if any error
ISNUMBER(value), ISTEXT(value), ISBLANK(value)  Type checks
```

---

## Array & Dynamic Functions {#array}

### Excel 365 Dynamic Arrays
```
FILTER(array, include, [if_empty])     Filter rows by condition
SORT(array, [sort_index], [sort_order])  Sort array
UNIQUE(array, [by_col], [exactly_once]) Unique values
SEQUENCE(rows, [cols], [start], [step])  Generate number sequence
RANDARRAY(rows, [cols], [min], [max])   Random number array
LET(name1, value1, ..., calculation)    Named variables in formula
LAMBDA(parameter, formula)              Custom functions (365+)
```

### Google Sheets Equivalents
```
FILTER()      — Same syntax as Excel
SORT()        — Same syntax as Excel
UNIQUE()      — Same syntax as Excel
QUERY()       — SQL-like querying (SHEETS ONLY, very powerful)
ARRAYFORMULA()  — Wrap formula to apply to entire range
IMPORTRANGE()   — Pull data from another spreadsheet (SHEETS ONLY)
GOOGLEFINANCE()  — Live financial data (SHEETS ONLY)
SPARKLINE()     — Inline mini charts (SHEETS ONLY)
```

---

## Excel vs Google Sheets Differences {#differences}

### Functions Only in Excel (Not in Google Sheets)
- XLOOKUP, XMATCH
- LET, LAMBDA
- STOCKHISTORY
- TEXTSPLIT (use SPLIT in Sheets)
- SEQUENCE (available in newer Sheets)

### Functions Only in Google Sheets (Not in Excel)
- QUERY (SQL-like, extremely powerful)
- IMPORTRANGE (cross-spreadsheet reference)
- GOOGLEFINANCE (live stock data)
- SPARKLINE (inline charts)
- ARRAYFORMULA (wrap for array behavior)
- IMPORTDATA, IMPORTHTML, IMPORTXML (web imports)

### Syntax Differences

| Feature | Excel | Google Sheets |
|---------|-------|---------------|
| Array formula entry | Ctrl+Shift+Enter (legacy) | ARRAYFORMULA() wrapper |
| Argument separator | `,` (US) or `;` (EU) | `,` (US) or `;` (EU) |
| Text split | TEXTSPLIT() | SPLIT() |
| Date system | 1900 (Jan 1, 1900 = 1) | 1899 (Dec 30, 1899 = 0) |
| Regular expressions | Not native | REGEXMATCH, REGEXEXTRACT, REGEXREPLACE |

### Behavioral Differences

1. **Array handling**: Sheets natively spills arrays; Excel requires dynamic array support (365+)
2. **Volatile functions**: Both recalculate RAND, NOW, TODAY on every change
3. **Maximum rows**: Excel: 1,048,576; Sheets: 10,000,000 (but slower)
4. **Maximum columns**: Excel: 16,384 (XFD); Sheets: 18,278
5. **File size**: Excel: limited by RAM; Sheets: 100 MB

### When to Use Which

**Use Excel for:**
- Complex financial models (DCF, LBO)
- Large datasets (100K+ rows)
- VBA macros and automation
- Power Query data transformation
- Investment banking deliverables

**Use Google Sheets for:**
- Real-time collaboration
- Simple tracking and dashboards
- Web data imports (IMPORTRANGE, GOOGLEFINANCE)
- Budget-conscious environments
- Quick sharing without file attachments

---

## Named Ranges & Best Practices {#naming}

### Why Use Named Ranges

```
Without: =B5*(1+$E$2)*(1-$E$3)
With:    =Revenue*(1+Growth_Rate)*(1-Churn_Rate)
```

### Naming Conventions

```
Good names:
  Revenue_Growth_Rate
  Tax_Rate_Federal
  Discount_Rate_WACC
  Assumption_Inflation

Bad names:
  x, temp, data, rate (too vague)
  Revenue Growth Rate (spaces not allowed)
  2024_Revenue (can't start with number)
```

### Named Range in openpyxl

```python
from openpyxl.workbook.defined_name import DefinedName

# Create named range
ref = "Assumptions!$B$5"
defn = DefinedName("Tax_Rate", attr_text=ref)
wb.defined_names.add(defn)

# Use in formula
ws['C10'] = '=Revenue * (1 - Tax_Rate)'
```

### Formula Readability Tips

1. Break complex formulas into intermediate cells with labels
2. Use named ranges for all assumptions
3. Add cell comments explaining non-obvious logic
4. Keep nesting depth ≤ 3 levels (use helper columns for more)
5. Document the purpose of each sheet in cell A1 or a TOC sheet
