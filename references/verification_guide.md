# Verification & Backtesting Guide

## Table of Contents
1. [Verification Philosophy](#philosophy)
2. [Shadow Calculation Pattern](#shadow-calc)
3. [Tolerance & Precision](#tolerance)
4. [Financial Verification](#financial)
5. [Engineering Verification](#engineering)
6. [Data Quality Checks](#data-quality)
7. [Property-Based Testing](#property-testing)
8. [Common Pitfalls](#pitfalls)

---

## Verification Philosophy {#philosophy}

Every spreadsheet calculation is independently verified by computing the same result in Python and comparing. This catches:

- Formula errors (wrong cell references, off-by-one ranges)
- Logic errors (wrong calculation sequence, missing steps)
- Rounding errors (floating point precision differences)
- Data errors (wrong input values, type mismatches)
- Structural errors (circular references, broken links)

The principle: **Two independent calculations arriving at the same result provides confidence. One calculation provides hope.**

---

## Shadow Calculation Pattern {#shadow-calc}

### Step-by-Step Process

```python
import pandas as pd
from openpyxl import load_workbook
import math
import json

def full_verification(filepath: str) -> dict:
    """Complete verification pipeline for any spreadsheet."""

    report = {
        'file': filepath,
        'checks': [],
        'summary': {'total': 0, 'passed': 0, 'failed': 0, 'warnings': 0}
    }

    # Phase 1: Structural checks
    structural = check_structure(filepath)
    report['checks'].extend(structural)

    # Phase 2: Formula error scan (via recalc.py)
    formula_errors = check_formula_errors(filepath)
    report['checks'].extend(formula_errors)

    # Phase 3: Value verification (shadow calculations)
    value_checks = verify_calculations(filepath)
    report['checks'].extend(value_checks)

    # Phase 4: Cross-reference checks
    xref_checks = verify_cross_references(filepath)
    report['checks'].extend(xref_checks)

    # Phase 5: Domain-specific checks
    domain_checks = verify_domain_rules(filepath)
    report['checks'].extend(domain_checks)

    # Summarize
    for check in report['checks']:
        report['summary']['total'] += 1
        if check['status'] == 'PASS':
            report['summary']['passed'] += 1
        elif check['status'] == 'FAIL':
            report['summary']['failed'] += 1
        else:
            report['summary']['warnings'] += 1

    return report
```

### Structural Checks

```python
def check_structure(filepath: str) -> list:
    checks = []
    wb = load_workbook(filepath, data_only=False)

    # Check for circular references (basic detection)
    # Build dependency graph from formulas
    deps = {}
    for ws in wb:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    refs = extract_cell_references(cell.value)
                    cell_addr = f"{ws.title}!{cell.coordinate}"
                    deps[cell_addr] = refs

    cycles = detect_cycles(deps)
    checks.append({
        'name': 'Circular reference check',
        'status': 'PASS' if not cycles else 'FAIL',
        'detail': f"Found {len(cycles)} circular references" if cycles else "No circular references"
    })

    # Check for empty formula cells
    empty_formulas = []
    wb_vals = load_workbook(filepath, data_only=True)
    for ws_name in wb.sheetnames:
        ws_f = wb[ws_name]
        ws_v = wb_vals[ws_name]
        for row in ws_f.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    val = ws_v[cell.coordinate].value
                    if val is None:
                        empty_formulas.append(f"{ws_name}!{cell.coordinate}")

    checks.append({
        'name': 'Empty formula check',
        'status': 'PASS' if not empty_formulas else 'WARNING',
        'detail': f"{len(empty_formulas)} formulas evaluate to empty" if empty_formulas else "All formulas have values"
    })

    return checks
```

### Value Verification

```python
def verify_sum_formulas(ws_formulas, ws_values, sheet_name: str) -> list:
    """Verify all SUM formulas independently."""
    checks = []

    for row in ws_formulas.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                formula = cell.value.upper()

                if formula.startswith('=SUM('):
                    # Parse the range
                    range_str = formula[5:-1]  # Remove =SUM( and )
                    # Calculate sum independently
                    python_sum = 0
                    for target_cell in ws_values[range_str]:
                        for c in target_cell:
                            if isinstance(c.value, (int, float)):
                                python_sum += c.value

                    excel_val = ws_values[cell.coordinate].value

                    match = verify_match(excel_val, python_sum)
                    checks.append({
                        'name': f'SUM verification: {sheet_name}!{cell.coordinate}',
                        'status': 'PASS' if match else 'FAIL',
                        'excel': excel_val,
                        'python': python_sum,
                        'formula': cell.value
                    })

    return checks
```

---

## Tolerance & Precision {#tolerance}

### Why Exact Equality Fails

```python
# This fails due to IEEE 754 floating point
0.1 + 0.1 + 0.1 == 0.3  # False!
# Result: 0.30000000000000004

# Always use tolerance-based comparison
math.isclose(0.1 + 0.1 + 0.1, 0.3)  # True
```

### Tolerance Levels by Domain

| Domain | Relative Tolerance | Absolute Tolerance | Notes |
|--------|-------------------|-------------------|-------|
| Financial (currency) | 1e-9 | 0.01 | To the cent |
| Financial (percentages) | 1e-6 | 0.0001 | 0.01% |
| Financial (multiples) | 1e-4 | 0.01 | 0.01x |
| Engineering (stress) | 1e-6 | depends on units | |
| Engineering (dimensions) | 1e-8 | depends on tolerance class | |
| Statistics (mean, std) | 1e-10 | — | Very tight |
| General counting | 0 | 0 | Exact match |

### Comparison Functions

```python
import math
from decimal import Decimal

def verify_match(excel_val, python_val, domain='general'):
    """Compare values with domain-appropriate tolerance."""
    if excel_val is None and python_val is None:
        return True
    if excel_val is None or python_val is None:
        return False
    if isinstance(excel_val, str) or isinstance(python_val, str):
        return str(excel_val) == str(python_val)

    tolerances = {
        'financial_currency': {'rel_tol': 1e-9, 'abs_tol': 0.01},
        'financial_pct': {'rel_tol': 1e-6, 'abs_tol': 0.0001},
        'engineering': {'rel_tol': 1e-6, 'abs_tol': 0},
        'statistics': {'rel_tol': 1e-10, 'abs_tol': 0},
        'general': {'rel_tol': 1e-9, 'abs_tol': 0.01},
        'exact': {'rel_tol': 0, 'abs_tol': 0},
    }

    tol = tolerances.get(domain, tolerances['general'])
    return math.isclose(float(excel_val), float(python_val), **tol)


def verify_financial(excel_val, python_val):
    """For currency values: match to the cent."""
    return verify_match(excel_val, python_val, 'financial_currency')
```

### When Excel and Python Disagree

Common causes and solutions:

1. **Excel truncates to 15 digits**: For very large numbers, expect small differences
2. **Rounding mode differences**: Excel uses "round half away from zero", Python uses "round half to even" (banker's rounding)
3. **Function implementation**: Some Excel functions use different algorithms than scipy/numpy
4. **Date system**: Excel 1900 date system vs Python datetime (1-day offset for dates before March 1, 1900)

---

## Financial Verification {#financial}

### Three-Statement Model Checks

```python
def verify_three_statement_model(filepath: str) -> list:
    checks = []

    wb = load_workbook(filepath, data_only=True)

    # 1. Balance Sheet Equation
    bs = wb['Balance Sheet']  # or appropriate sheet name
    total_assets = get_total(bs, 'Total Assets')
    total_liabilities = get_total(bs, 'Total Liabilities')
    total_equity = get_total(bs, 'Total Equity')

    for period in range(num_periods):
        a = total_assets[period]
        l = total_liabilities[period]
        e = total_equity[period]
        checks.append({
            'name': f'Balance Sheet Eq (Period {period})',
            'status': 'PASS' if math.isclose(a, l + e, abs_tol=0.01) else 'FAIL',
            'detail': f'Assets={a:.2f}, L+E={l+e:.2f}, diff={a-l-e:.2f}'
        })

    # 2. Cash Flow Reconciliation
    cf = wb['Cash Flow']
    for period in range(1, num_periods):
        opening = get_value(bs, 'Cash', period - 1)
        net_change = get_value(cf, 'Net Change in Cash', period)
        closing = get_value(bs, 'Cash', period)
        checks.append({
            'name': f'Cash Reconciliation (Period {period})',
            'status': 'PASS' if math.isclose(opening + net_change, closing, abs_tol=0.01) else 'FAIL',
            'detail': f'Opening={opening:.2f} + Change={net_change:.2f} = {opening+net_change:.2f}, Closing={closing:.2f}'
        })

    # 3. Net Income Linkage
    pl = wb['Income Statement']
    for period in range(num_periods):
        pl_ni = get_value(pl, 'Net Income', period)
        cf_ni = get_value(cf, 'Net Income', period)
        checks.append({
            'name': f'Net Income Linkage (Period {period})',
            'status': 'PASS' if math.isclose(pl_ni, cf_ni, abs_tol=0.01) else 'FAIL',
            'detail': f'P&L={pl_ni:.2f}, CF={cf_ni:.2f}'
        })

    return checks
```

### DCF Verification

```python
import numpy as np

def verify_dcf(cash_flows, discount_rate, terminal_value, terminal_year):
    """Verify DCF calculation independently."""

    # NPV of projected cash flows
    python_npv = sum(cf / (1 + discount_rate)**t
                     for t, cf in enumerate(cash_flows, 1))

    # Terminal value discounted
    python_tv_pv = terminal_value / (1 + discount_rate)**terminal_year

    python_ev = python_npv + python_tv_pv

    return python_ev

def verify_irr(cash_flows):
    """Verify IRR using numpy."""
    return np.irr(cash_flows)

def verify_xirr(cash_flows, dates):
    """Verify XIRR using scipy optimization."""
    from scipy.optimize import brentq
    from datetime import datetime

    def xnpv(rate):
        d0 = dates[0]
        return sum(cf / (1 + rate)**((d - d0).days / 365.25)
                   for cf, d in zip(cash_flows, dates))

    return brentq(xnpv, -0.99, 10.0)
```

### Sensitivity Verification

```python
def verify_sensitivity_table(ws, base_cell, input_cells, expected_outputs):
    """Verify that sensitivity table values are consistent."""
    checks = []

    for i, (input_val, expected_output) in enumerate(zip(input_cells, expected_outputs)):
        # Each sensitivity point should follow the formula:
        # output = f(input_val) where f is the model function
        python_output = model_function(input_val)
        excel_output = expected_outputs[i]

        checks.append({
            'name': f'Sensitivity point {i}: input={input_val}',
            'status': 'PASS' if verify_financial(excel_output, python_output) else 'FAIL',
            'excel': excel_output,
            'python': python_output
        })

    return checks
```

---

## Engineering Verification {#engineering}

### Unit Consistency Check

```python
from pint import UnitRegistry
ureg = UnitRegistry()

def verify_engineering_calc(formula_description, inputs, expected_output, expected_units):
    """Verify an engineering calculation with unit tracking."""

    # Recreate calculation with units
    result = compute_with_units(inputs)

    # Check magnitude
    magnitude_match = math.isclose(
        result.magnitude,
        expected_output,
        rel_tol=1e-6
    )

    # Check units
    units_match = result.units == expected_units

    return {
        'name': formula_description,
        'magnitude_match': magnitude_match,
        'units_match': units_match,
        'computed': f"{result:.4f}",
        'expected': f"{expected_output} {expected_units}"
    }
```

### Boundary Condition Checks

```python
def verify_engineering_boundaries(results: dict) -> list:
    """Check physical reasonableness of engineering results."""
    checks = []

    # Stress must be positive for tension, reasonable magnitude
    if 'stress' in results:
        checks.append({
            'name': 'Stress reasonableness',
            'status': 'PASS' if 0 < results['stress'] < 1e6 else 'WARNING',
            'detail': f"Stress = {results['stress']:.2f} MPa"
        })

    # Factor of safety must be > 1.0
    if 'factor_of_safety' in results:
        checks.append({
            'name': 'Factor of safety > 1.0',
            'status': 'PASS' if results['factor_of_safety'] > 1.0 else 'FAIL',
            'detail': f"FoS = {results['factor_of_safety']:.2f}"
        })

    # Efficiency must be 0-100%
    if 'efficiency' in results:
        checks.append({
            'name': 'Efficiency in valid range',
            'status': 'PASS' if 0 <= results['efficiency'] <= 1.0 else 'FAIL',
            'detail': f"η = {results['efficiency']*100:.1f}%"
        })

    return checks
```

---

## Data Quality Checks {#data-quality}

### Schema Validation

```python
def validate_schema(df: pd.DataFrame, schema: dict) -> list:
    """Validate DataFrame against expected schema."""
    checks = []

    for col_name, rules in schema.items():
        if col_name not in df.columns:
            checks.append({'name': f'Column exists: {col_name}', 'status': 'FAIL'})
            continue

        checks.append({'name': f'Column exists: {col_name}', 'status': 'PASS'})

        # Type check
        if 'dtype' in rules:
            actual_type = df[col_name].dtype
            checks.append({
                'name': f'Type check: {col_name}',
                'status': 'PASS' if str(actual_type).startswith(rules['dtype']) else 'WARNING',
                'detail': f'Expected {rules["dtype"]}, got {actual_type}'
            })

        # Null check
        if rules.get('required', False):
            null_count = df[col_name].isnull().sum()
            checks.append({
                'name': f'No nulls: {col_name}',
                'status': 'PASS' if null_count == 0 else 'FAIL',
                'detail': f'{null_count} null values found'
            })

        # Range check
        if 'min' in rules or 'max' in rules:
            min_val = df[col_name].min()
            max_val = df[col_name].max()
            in_range = True
            if 'min' in rules and min_val < rules['min']:
                in_range = False
            if 'max' in rules and max_val > rules['max']:
                in_range = False
            checks.append({
                'name': f'Range check: {col_name}',
                'status': 'PASS' if in_range else 'FAIL',
                'detail': f'Range [{min_val}, {max_val}], expected [{rules.get("min","")}, {rules.get("max","")}]'
            })

    return checks
```

### Completeness & Consistency

```python
def check_data_quality(df: pd.DataFrame) -> list:
    checks = []

    # Completeness
    for col in df.columns:
        completeness = (1 - df[col].isnull().sum() / len(df)) * 100
        checks.append({
            'name': f'Completeness: {col}',
            'status': 'PASS' if completeness >= 95 else 'WARNING' if completeness >= 80 else 'FAIL',
            'detail': f'{completeness:.1f}% complete'
        })

    # Duplicate detection
    dupes = df.duplicated().sum()
    checks.append({
        'name': 'Duplicate rows',
        'status': 'PASS' if dupes == 0 else 'WARNING',
        'detail': f'{dupes} duplicate rows'
    })

    # Outlier detection (IQR method)
    for col in df.select_dtypes(include='number').columns:
        Q1 = df[col].quantile(0.25)
        Q3 = df[col].quantile(0.75)
        IQR = Q3 - Q1
        outliers = ((df[col] < Q1 - 1.5*IQR) | (df[col] > Q3 + 1.5*IQR)).sum()
        checks.append({
            'name': f'Outliers: {col}',
            'status': 'PASS' if outliers == 0 else 'WARNING',
            'detail': f'{outliers} outliers detected'
        })

    return checks
```

---

## Property-Based Testing {#property-testing}

### Using Hypothesis for Edge Case Discovery

```python
from hypothesis import given, strategies as st

@given(
    revenue=st.floats(min_value=0, max_value=1e9, allow_nan=False),
    cogs=st.floats(min_value=0, max_value=1e9, allow_nan=False)
)
def test_gross_margin_always_valid(revenue, cogs):
    """Gross margin should always be between -inf and 1.0 when revenue > 0."""
    if revenue > 0:
        margin = (revenue - cogs) / revenue
        assert margin <= 1.0  # Can't have > 100% margin
        # COGS > Revenue means negative margin (valid but unusual)

@given(
    values=st.lists(st.floats(min_value=-1e6, max_value=1e6, allow_nan=False),
                    min_size=1, max_size=100)
)
def test_sum_matches_python(values):
    """Excel SUM should always match Python sum."""
    excel_sum = sum(values)  # Simulates Excel SUM
    python_sum = sum(values)
    assert math.isclose(excel_sum, python_sum, rel_tol=1e-10)
```

---

## Common Pitfalls {#pitfalls}

### Excel-Specific Precision Issues

1. **15-digit limit**: Excel stores max 15 significant digits. Numbers like 123456789012345**6** become 123456789012345**0**
2. **Small number subtraction**: (1e15 + 1) - 1e15 = 0 in Excel (should be 1)
3. **ROUND function**: Uses "round half away from zero", Python uses "round half to even"
4. **Date arithmetic**: Excel's 1900 date system has the Lotus 1-2-3 bug (Feb 29, 1900 exists in Excel but not in reality)

### Verification Anti-Patterns (DON'T DO)

```python
# WRONG: Using exact equality for floats
assert excel_value == python_value  # Will fail on 0.1+0.2

# WRONG: Comparing formatted strings
assert f"{excel_value:.2f}" == f"{python_value:.2f}"  # Hides real differences

# WRONG: Ignoring None/NaN differences
# Excel empty cell = None, Python might be NaN — these are different

# WRONG: Skipping verification "because the formula looks right"
# Visual inspection misses 90%+ of errors
```

### Verification Best Practices (DO)

```python
# RIGHT: Tolerance-based comparison
assert math.isclose(excel_value, python_value, rel_tol=1e-9, abs_tol=0.01)

# RIGHT: Handle None/NaN explicitly
if excel_val is None:
    assert python_val is None or math.isnan(python_val)

# RIGHT: Log everything for audit trail
log.info(f"Cell {addr}: Excel={excel_val}, Python={python_val}, Match={match}")

# RIGHT: Test with extreme values
test_cases = [0, -1, 1e-10, 1e15, float('inf')]
```
