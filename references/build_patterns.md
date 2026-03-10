# Build Patterns & Code Reference

## Table of Contents
1. [Workbook Creation](#creation)
2. [Formatting Patterns](#formatting)
3. [Chart Creation](#charts)
4. [Data Import](#import)
5. [Conditional Formatting](#conditional)
6. [Data Validation](#validation)
7. [Pivot Tables & Summaries](#pivot)
8. [Chart Gallery](#chart-gallery)
9. [Sensitivity Analysis](#sensitivity)
10. [Google Sheets Compatibility](#sheets)
11. [Large Dataset Handling](#large-data)
12. [Common Templates](#templates)

---

## Workbook Creation {#creation}

### Basic Structure

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers

wb = Workbook()
ws = wb.active
ws.title = "Summary"

# Standard styles
BLUE_INPUT = Font(name='Arial', size=10, color='0000FF')
BLACK_FORMULA = Font(name='Arial', size=10, color='000000')
GREEN_LINK = Font(name='Arial', size=10, color='008000')
HEADER_FONT = Font(name='Arial', size=11, bold=True)
HEADER_FILL = PatternFill('solid', fgColor='D9E1F2')
YELLOW_BG = PatternFill('solid', fgColor='FFFF00')
THIN_BORDER = Border(bottom=Side(style='thin'))
```

### Adding Sheets with Structure

```python
# Create sheets in logical order
ws_assumptions = wb.create_sheet("Assumptions", 0)
ws_data = wb.create_sheet("Data", 1)
ws_calcs = wb.create_sheet("Calculations", 2)
ws_output = wb.create_sheet("Output", 3)

# Remove default sheet if not needed
if "Sheet" in wb.sheetnames:
    del wb["Sheet"]
```

### Column Width Auto-Fit (Approximate)

```python
def auto_fit_columns(ws, min_width=8, max_width=50):
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)
```

### Freeze Panes

```python
# Freeze header row
ws.freeze_panes = 'A2'

# Freeze header row AND first column
ws.freeze_panes = 'B2'
```

---

## Formatting Patterns {#formatting}

### Financial Model Number Formats

```python
FMT_CURRENCY = '$#,##0;($#,##0);"-"'
FMT_CURRENCY_MM = '$#,##0.0;($#,##0.0);"-"'  # Millions with 1 decimal
FMT_PERCENT = '0.0%'
FMT_MULTIPLE = '0.0"x"'
FMT_YEAR = '@'  # Text, prevents 2024 → 2,024
FMT_INTEGER = '#,##0;(#,##0);"-"'
FMT_DECIMAL_2 = '#,##0.00;(#,##0.00);"-"'
FMT_DATE = 'YYYY-MM-DD'
FMT_ACCOUNTING = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
```

### Apply Formatting to Range

```python
def format_range(ws, cell_range, font=None, fill=None, number_format=None, alignment=None):
    for row in ws[cell_range]:
        for cell in row:
            if font: cell.font = font
            if fill: cell.fill = fill
            if number_format: cell.number_format = number_format
            if alignment: cell.alignment = alignment
```

### Header Row Formatting

```python
def format_headers(ws, row=1, start_col=1, end_col=10):
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = Border(bottom=Side(style='medium'))
```

---

## Chart Creation {#charts}

### Bar Chart

```python
from openpyxl.chart import BarChart, Reference

chart = BarChart()
chart.type = "col"
chart.title = "Revenue by Quarter"
chart.y_axis.title = "Revenue ($)"
chart.x_axis.title = "Quarter"
chart.style = 10

data = Reference(ws, min_col=2, min_row=1, max_row=5, max_col=2)
cats = Reference(ws, min_col=1, min_row=2, max_row=5)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4

ws.add_chart(chart, "E2")  # Position
chart.width = 15  # inches
chart.height = 10
```

### Line Chart

```python
from openpyxl.chart import LineChart

chart = LineChart()
chart.title = "Revenue Trend"
chart.y_axis.title = "Revenue"
chart.style = 10

data = Reference(ws, min_col=2, min_row=1, max_row=13, max_col=3)
cats = Reference(ws, min_col=1, min_row=2, max_row=13)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "E2")
```

### Pie Chart

```python
from openpyxl.chart import PieChart

chart = PieChart()
chart.title = "Revenue by Segment"

data = Reference(ws, min_col=2, min_row=1, max_row=6)
cats = Reference(ws, min_col=1, min_row=2, max_row=6)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws.add_chart(chart, "E2")
```

---

## Data Validation {#validation}

### Dropdown List from Range

```python
from openpyxl.worksheet.datavalidation import DataValidation

# Create dropdown from a named range
dv = DataValidation(type="list", formula1="$G$2:$G$10", allow_blank=True)
dv.error = 'Please select from list'
dv.errorTitle = 'Invalid Entry'
ws.add_data_validation(dv)
dv.add('B2:B100')  # Apply to column B

# Or inline formula
dv2 = DataValidation(type="list", formula1='"Option1,Option2,Option3"')
dv2.add('C2:C100')
ws.add_data_validation(dv2)
```

### Number Range Validation (Min/Max)

```python
from openpyxl.worksheet.datavalidation import DataValidation

# Whole numbers between 1 and 100
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="100")
dv.error = 'Must be between 1 and 100'
dv.errorTitle = 'Invalid Number'
ws.add_data_validation(dv)
dv.add('D2:D100')

# Decimal numbers ≥ 0 and ≤ 1 (useful for percentages as decimals)
dv_decimal = DataValidation(type="decimal", operator="between", formula1="0", formula2="1")
ws.add_data_validation(dv_decimal)
dv_decimal.add('E2:E100')
```

### Date Range Validation

```python
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime

# Dates between 2024-01-01 and 2024-12-31
dv = DataValidation(type="date", operator="between",
                    formula1="2024-01-01", formula2="2024-12-31")
dv.error = 'Date must be in 2024'
dv.errorTitle = 'Invalid Date'
ws.add_data_validation(dv)
dv.add('F2:F100')

# Dates greater than today (requires sheet formula reference)
# Use a cell with TODAY() and reference it
ws['H1'] = "=TODAY()"
dv_future = DataValidation(type="date", operator="greaterThan", formula1="$H$1")
ws.add_data_validation(dv_future)
dv_future.add('I2:I100')
```

### Custom Formula Validation

```python
from openpyxl.worksheet.datavalidation import DataValidation

# Value must be less than the value in column C on same row
dv = DataValidation(type="custom", formula1="=D2<C2", allow_blank=False)
dv.error = 'Value must be less than column C'
dv.prompt = 'Enter a value smaller than the limit'
dv.promptTitle = 'Input Required'
ws.add_data_validation(dv)
dv.add('D2:D100')

# Value must be unique (no duplicates in range)
# Note: This is complex; simpler to use a helper column with COUNTIF
dv_unique = DataValidation(type="custom", formula1="=COUNTIF($A$2:$A$100,A2)=1")
ws.add_data_validation(dv_unique)
dv_unique.add('A2:A100')
```

### Input Message and Error Alert Customization

```python
from openpyxl.worksheet.datavalidation import DataValidation

dv = DataValidation(
    type="list",
    formula1="$G$2:$G$10",
    allow_blank=False,
    showInputMessage=True,
    prompt="Select a valid department",
    promptTitle="Department Selection",
    showErrorMessage=True,
    error="Invalid selection. Please choose from the list.",
    errorTitle="Error: Invalid Department"
)
ws.add_data_validation(dv)
dv.add('B2:B100')
```

---

## Pivot Tables & Summaries {#pivot}

### Pandas Pivot Table Approach (Recommended)

**Note:** openpyxl has limited built-in pivot table support. The pandas approach creates static summary tables that function like pivot tables with proper formatting.

```python
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Create pivot from pandas
df = pd.read_excel('data.xlsx', sheet_name='Data')

# Pivot with aggregation
pivot = df.pivot_table(
    values='Amount',
    index='Department',
    columns='Month',
    aggfunc='sum',
    margins=True  # Adds grand total row
)

# Write to new sheet
with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
    pivot.to_excel(writer, sheet_name='Pivot')

# Format the pivot sheet
wb = load_workbook('output.xlsx')
ws = wb['Pivot']

# Format header row
for cell in ws[1]:
    if cell.value:
        cell.font = Font(bold=True, size=11)
        cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        cell.alignment = Alignment(horizontal='center')

# Format index column
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
    for cell in row:
        cell.font = Font(bold=True)

# Format numbers as currency
for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2):
    for cell in row:
        if cell.value and isinstance(cell.value, (int, float)):
            cell.number_format = '$#,##0'

wb.save('output.xlsx')
```

### Manual Summary Table with Subtotals

```python
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Create summary by grouping
summary = df.groupby(['Department']).agg({
    'Amount': ['sum', 'count', 'mean']
}).round(2)

summary.columns = ['Total', 'Count', 'Average']
summary['% of Grand Total'] = (summary['Total'] / summary['Total'].sum() * 100).round(1)

# Add grand total row
grand_total = pd.DataFrame({
    'Total': [summary['Total'].sum()],
    'Count': [summary['Count'].sum()],
    'Average': [summary['Amount'].mean()],
    '% of Grand Total': [100.0]
}, index=['GRAND TOTAL'])

summary = pd.concat([summary, grand_total])

# Write to Excel
summary.to_excel('summary.xlsx', sheet_name='Summary')

# Format using openpyxl
wb = load_workbook('summary.xlsx')
ws = wb.active

# Header formatting
header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
header_font = Font(color='FFFFFF', bold=True, size=11)
for cell in ws[1]:
    if cell.value:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')

# Grand total row formatting (last row)
gt_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
gt_font = Font(bold=True)
for cell in ws[ws.max_row]:
    cell.fill = gt_fill
    cell.font = gt_font

# Number formatting
for row in ws.iter_rows(min_row=2, max_col=4):
    for cell in row:
        if cell.column in [2, 3, 4]:  # Total, Count, Average
            cell.number_format = '#,##0.00'
        elif cell.column == 5:  # Percentage
            cell.number_format = '0.0"%"'

wb.save('summary.xlsx')
```

### Adding Subtotals and Grand Totals in Code

```python
# Using SUBTOTAL function in openpyxl formulas
from openpyxl.utils import get_column_letter

def add_subtotals(ws, data_start_row=2, data_end_row=100, summary_col=2):
    """Add subtotal formulas for grouped data"""
    # SUBTOTAL function: 9=SUM, 103=SUM(ignore hidden), etc.

    # Subtotal row for first group
    ws[f'A101'] = 'Subtotal'
    ws[f'{get_column_letter(summary_col)}101'] = f'=SUBTOTAL(9,{get_column_letter(summary_col)}{data_start_row}:{get_column_letter(summary_col)}{data_end_row})'

    # Grand total row
    ws['A102'] = 'GRAND TOTAL'
    ws[f'{get_column_letter(summary_col)}102'] = f'=SUM({get_column_letter(summary_col)}101)'

    # Format subtotal/grand total rows
    for cell in ws[101]:
        cell.font = Font(bold=True)
    for cell in ws[102]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
```

---

## Chart Gallery {#chart-gallery}

### Waterfall Chart (Using Stacked Bar Technique)

**Note:** openpyxl doesn't natively support waterfall charts, so we build one using stacked bars with invisible connector bars.

```python
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as XLImage

# Prepare data structure:
# Col A: Category (Opening Balance, Sales, Returns, Closing Balance)
# Col B: Value (cumulative position - used for invisible base)
# Col C: Amount (height of visible bar)

# Example layout in worksheet:
# A           B    C
# Category    Pos  Amount
# Opening     0    100
# Sales       100  50
# Returns     150  -20
# Closing     130  -

# Create stacked bar chart
chart = BarChart()
chart.type = "col"
chart.title = "Waterfall Chart"
chart.grouping = "percentStacked"  # Use standard stacked

# Data references
pos = Reference(ws, min_col=2, min_row=1, max_row=5)  # Invisible base (Y values)
amount = Reference(ws, min_col=3, min_row=1, max_row=5)  # Visible amounts
cats = Reference(ws, min_col=1, min_row=2, max_row=5)

chart.add_data(pos, titles_from_data=True)
chart.add_data(amount, titles_from_data=True)
chart.set_categories(cats)

# Format: hide the position series (make it invisible)
chart.series[0].graphicalProperties.solidFill = "F0F0F0"  # Match background
chart.dataLabels = DataLabelList()
chart.dataLabels.showVal = True

ws.add_chart(chart, "E2")
chart.width = 15
chart.height = 10
```

### Scatter Plot with Trendline

```python
from openpyxl.chart import ScatterChart, Reference
from openpyxl.chart.trendline import Trendline

chart = ScatterChart()
chart.title = "Scatter: Price vs Volume"
chart.x_axis.title = "Volume (units)"
chart.y_axis.title = "Price ($)"
chart.style = 10

# Data: X values, Y values
xvalues = Reference(ws, min_col=1, min_row=2, max_row=50)
values = Reference(ws, min_col=2, min_row=1, max_row=50)

chart.add_data(values, titles_from_data=True)
chart.set_categories(xvalues)

# Add trendline (linear)
trendline = Trendline()
trendline.type = "linear"  # or "exp", "log", "power", "poly"
trendline.degree = 1  # For polynomial: 2, 3, etc.
trendline.dispEq = True  # Show equation
trendline.dispR2 = True  # Show R² value

chart.series[0].trendline = trendline

ws.add_chart(chart, "E2")
```

### Combo Chart (Bar + Line on Dual Axes)

```python
from openpyxl.chart import BarChart, LineChart, Reference

# Create bar chart first
bar = BarChart()
bar.type = "col"
bar.title = "Sales vs Margin %"
bar.y_axis.title = "Sales ($)"

data_sales = Reference(ws, min_col=2, min_row=1, max_row=13)
cats = Reference(ws, min_col=1, min_row=2, max_row=13)
bar.add_data(data_sales, titles_from_data=True)
bar.set_categories(cats)

# Create line chart for secondary axis
line = LineChart()
line.y_axis.title = "Margin %"
line.y_axis.crosses = "max"  # Secondary axis

data_margin = Reference(ws, min_col=3, min_row=1, max_row=13)
line.add_data(data_margin, titles_from_data=True)

# Combine charts
bar += line

ws.add_chart(bar, "E2")
bar.width = 16
bar.height = 10
```

### Area Chart

```python
from openpyxl.chart import AreaChart, Reference

chart = AreaChart()
chart.title = "Revenue Breakdown by Region"
chart.style = 10
chart.grouping = "percentStacked"  # or "standard", "stackedPercent"

data = Reference(ws, min_col=2, min_row=1, max_row=13, max_col=4)
cats = Reference(ws, min_col=1, min_row=2, max_row=13)

chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Customize fills
chart.series[0].graphicalProperties.solidFill = "FF6B6B"  # Series 1 color
chart.series[1].graphicalProperties.solidFill = "4ECDC4"  # Series 2 color

ws.add_chart(chart, "E2")
```

### Sparkline-Like Inline Charts (Using Data Bars)

**Note:** openpyxl doesn't support true sparklines, but conditional formatting data bars provide a similar visual effect.

```python
from openpyxl.formatting.rule import DataBarRule

# Create mini bar chart using conditional data bar
# Useful for inline trend indicators in dashboards

# Apply data bar to column (e.g., monthly trend in a summary table)
rule = DataBarRule(
    start_type='min',
    end_type='max',
    color='638EC6',
    showValue=True  # Show numbers AND bars
)

ws.conditional_formatting.add('B2:B100', rule)

# For a range like B2:M2 (monthly data across row), apply per-row:
for row in range(2, 50):
    rule_row = DataBarRule(
        start_type='min',
        end_type='max',
        color='92D050'
    )
    ws.conditional_formatting.add(f'B{row}:M{row}', rule_row)
```

---

## Sensitivity Analysis / Data Table {#sensitivity}

### One-Variable Data Table

**Setup:** Input parameter varies in rows or columns; formula result updates accordingly.

```python
# Example: How does NPV change with different discount rates?

# In worksheet:
# A1: "Discount Rate"  B1: "NPV"
# A2: 0.05             B2: =NPV(A2, cashflows...)
# A3: 0.06
# A4: 0.07
# A5: 0.08
# ... etc

# Data table in Excel would be:
# (empty)     5%      6%      7%      8%
# NPV         val1    val2    val3    val4

# Build programmatically:
from openpyxl.utils import get_column_letter

ws['A1'] = "Discount Rate"
ws['B1'] = "NPV"

rates = [0.05, 0.06, 0.07, 0.08, 0.09, 0.10]
for idx, rate in enumerate(rates, start=2):
    ws[f'A{idx}'] = rate
    # Formula references input cell and calculates NPV
    ws[f'B{idx}'] = f'=NPV(A{idx},CashFlows!$B$2:$B$10)'

# Format
for row in range(1, len(rates) + 2):
    ws[f'A{row}'].number_format = '0.0%'
    ws[f'B{row}'].number_format = '$#,##0'
```

### Two-Variable Data Table

**Setup:** Two input parameters vary (rows × columns); result at intersection.

```python
# Example: Sensitivity table: Discount Rate (rows) × Growth Rate (columns)

# Layout in worksheet:
#           3%      4%      5%      6%
# 5%        NPV1    NPV2    NPV3    NPV4
# 6%        NPV5    NPV6    NPV7    NPV8
# 7%        NPV9    NPV10   NPV11   NPV12
# 8%        NPV13   NPV14   NPV15   NPV16

# Build two-variable table:
discount_rates = [0.05, 0.06, 0.07, 0.08]
growth_rates = [0.03, 0.04, 0.05, 0.06]

# Header row: Growth rates
for idx, gr in enumerate(growth_rates, start=2):
    ws.cell(row=1, column=idx).value = gr
    ws.cell(row=1, column=idx).number_format = '0.0%'

# Header col: Discount rates
for idx, dr in enumerate(discount_rates, start=2):
    ws.cell(row=idx, column=1).value = dr
    ws.cell(row=idx, column=1).number_format = '0.0%'

# Values: Formula at intersection
for row_idx, dr in enumerate(discount_rates, start=2):
    for col_idx, gr in enumerate(growth_rates, start=2):
        cell = ws.cell(row=row_idx, column=col_idx)
        # Reference input cells: one for discount, one for growth
        cell.value = f'=NPV(A{row_idx}, CashFlows!B:B) * (1 + B$1)'
        cell.number_format = '$#,##0'

# Format table
for cell in ws[1]:
    if cell.value:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
```

### Scenario Comparison Layout

**Setup:** Compare base case, upside, and downside scenarios side-by-side.

```python
# Example: 3-scenario projection

# Layout:
# Item              Base    Upside  Downside
# Revenue Growth    5%      10%     2%
# Operating Margin  20%     25%     18%
# NPV (Result)      $100M   $150M   $60M

# Build in code:
scenarios = {
    'Base': {'rev_growth': 0.05, 'op_margin': 0.20},
    'Upside': {'rev_growth': 0.10, 'op_margin': 0.25},
    'Downside': {'rev_growth': 0.02, 'op_margin': 0.18}
}

ws['A1'] = "Item"
ws['B1'] = "Base"
ws['C1'] = "Upside"
ws['D1'] = "Downside"

# Assumptions
row = 2
for key in ['rev_growth', 'op_margin']:
    ws[f'A{row}'] = key.replace('_', ' ').title()
    for col_idx, (scenario_name, values) in enumerate(scenarios.items(), start=2):
        cell = ws.cell(row=row, column=col_idx)
        cell.value = values[key]
        cell.number_format = '0.0%'
    row += 1

# Results (formulas that reference assumptions)
ws[f'A{row}'] = "NPV Result"
for col_idx, (scenario_name, _) in enumerate(scenarios.items(), start=2):
    cell = ws.cell(row=row, column=col_idx)
    # Formula using scenario's assumption cells
    cell.value = f'=NPV({get_column_letter(col_idx)}2, Cashflows!$A$1:$A$10)'
    cell.number_format = '$#,##0'

# Format headers
for cell in ws[1]:
    if cell.value:
        cell.font = Font(bold=True, size=11)
        cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        cell.font = Font(bold=True, color='FFFFFF')
```

---

## Google Sheets Compatibility {#sheets}

### When to Recommend Sheets vs Excel

**Use Excel (.xlsx) when:**
- Heavy formulas, VBA macros, or complex calculations
- Pivot tables are essential
- File will be large (100K+ rows)
- Advanced conditional formatting or chart types needed
- One-time analysis or client deliverable with strict formatting

**Use Google Sheets when:**
- Real-time collaboration is required
- File needs to be shared and edited by multiple people
- Simple dashboards or forms integration
- Mobile access is needed
- Integration with Google Forms, Data Studio, or other Google tools

**Hybrid approach:**
- Build in Excel, publish summary to Sheets
- Use Sheets for data collection, analyze in Excel

### Key Differences: Building Sheets vs Excel

#### Differences in Formula Syntax

| Concept | Excel | Google Sheets |
|---------|-------|---------------|
| Colon syntax | `A1:A10` | `A1:A10` (same) |
| Sheet reference | `Sheet1!A1` | `'Sheet 1'!A1` (quotes if space) |
| Array formulas | `Ctrl+Shift+Enter` | `=ARRAYFORMULA()` wrapper |
| SUMIF | `=SUMIF(range, criteria, sum_range)` | Same |
| COUNTIF | `=COUNTIF(range, criteria)` | Same |
| VLOOKUP | `=VLOOKUP(lookup_value, table_array, col_index, FALSE)` | Same (use FALSE not 0) |
| UNIQUE | Not native | `=UNIQUE(range)` (native) |
| FILTER | Not native | `=FILTER(range, condition)` (native) |
| INDEX/MATCH | Same | Same |
| IF multiple conditions | `=IF(AND(cond1, cond2), val1, val2)` | Same, but ARRAYFORMULA for arrays |

#### Sheets-specific Functions (No Excel Equivalent)

```
=QUERY()       - SQL-like queries on ranges
=FILTER()      - Dynamic filtering
=UNIQUE()      - Get unique values
=SEQUENCE()    - Generate sequences
=REGEXMATCH()  - Regex matching
=IMAGEURL()    - Embed images from URL
=SPARKLINE()   - Native sparklines (not available in openpyxl)
```

#### Approach with gspread (Python Library for Google Sheets)

```python
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Authenticate with service account
scope = ['https://spreadsheets.google.com/feeds']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# Open sheet
sheet = client.open('My Sheet').sheet1

# Read data
all_cells = sheet.get_all_values()

# Write data
sheet.update_cells([
    gspread.Cell(row=1, col=1, value='Header'),
    gspread.Cell(row=2, col=1, value='Data')
])

# Append rows
sheet.append_row(['Value1', 'Value2', 'Value3'])

# Format cells (limited compared to openpyxl)
fmt = gspread.formatting.CellFormat(
    backgroundColor=gspread.utils.color(1, 1, 1),  # RGB values
    textFormat={'bold': True}
)
sheet.format([('A1:C1', fmt)])

# Add data validation (Sheets supports this)
rule = gspread.worksheet.DataValidationRule(
    type='DROPDOWN',
    inputMessage='Select from list',
    strict=True,
    values=['Option1', 'Option2']
)
sheet.data_validation_add('B2:B100', rule)
```

### Formula Syntax Differences to Watch For

1. **References to other sheets:**
   - Excel: `=Sheet1!A1` or `='Sheet 1'!A1` (if space)
   - Sheets: `=Sheet1!A1` or `='Sheet 1'!A1` (same rules)

2. **Array formulas:**
   - Excel: `{=SUM(A1:A10 * B1:B10)}` (entered with Ctrl+Shift+Enter)
   - Sheets: `=ARRAYFORMULA(SUM(A1:A10 * B1:B10))` (explicit wrapper)

3. **Wildcards in COUNTIF/SUMIF:**
   - Excel: `=COUNTIF(range, "prefix*")`
   - Sheets: Same, but REGEX functions are alternative: `=SUMPRODUCT((REGEXMATCH(range,"^prefix")*values))`

4. **Date functions:**
   - Excel: `=TODAY()`, `=NOW()` (same)
   - Sheets: Same, but TODATE() exists for conversions

5. **Text functions:**
   - Excel: `=CONCATENATE()` or `&` or `=CONCAT()`
   - Sheets: `=CONCATENATE()`, `&`, `=CONCAT()` (all work), also `=TEXTJOIN()`

### Notes on Primary Output Format

**Primary output is always `.xlsx` (Excel format)** because:
- Better support for complex formulas and formatting
- openpyxl is more mature and feature-rich than gspread
- .xlsx is universally compatible
- Easier to handle large datasets and calculations

**Document Sheets compatibility:**
- Include notes in README: "Excel file created with openpyxl. For Google Sheets collaboration, import the .xlsx file to Sheets (File → Import → Upload)."
- List formula adjustments needed (ARRAYFORMULA, UNIQUE, FILTER if sheets-specific).
- Note data validation compatibility: Both support similar validation, but formatting options differ slightly.
- Provide a "Sheets Export Checklist" if users need to move to Sheets:
  - Remove unsupported functions (VBA, certain complex array formulas)
  - Re-apply conditional formatting (rules export but may need tweaking)
  - Verify pivot tables (Sheets has different pivot mechanics)
  - Test chart compatibility (some chart types may need recreation)

---

## Data Import {#import}

### CSV Import

```python
import pandas as pd

# Auto-detect encoding and delimiter
df = pd.read_csv('data.csv', encoding='utf-8-sig')
# If encoding fails, try: encoding='latin-1' or encoding='cp1252'

# Write to Excel preserving types
with pd.ExcelWriter('output.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='Data')
```

### JSON Import

```python
import json
df = pd.json_normalize(json.load(open('data.json')))
df.to_excel('output.xlsx', index=False)
```

### Multiple CSV Files into Sheets

```python
with pd.ExcelWriter('combined.xlsx', engine='openpyxl') as writer:
    for csv_file in csv_files:
        df = pd.read_csv(csv_file)
        sheet_name = Path(csv_file).stem[:31]  # Excel 31-char limit
        df.to_excel(writer, index=False, sheet_name=sheet_name)
```

### Existing Excel File Reading

```python
# Read all sheets
all_sheets = pd.read_excel('input.xlsx', sheet_name=None)
for name, df in all_sheets.items():
    print(f"Sheet: {name}, Shape: {df.shape}")

# Read with types
df = pd.read_excel('input.xlsx', dtype={'ID': str, 'Amount': float})
```

---

## Conditional Formatting {#conditional}

### Traffic Light (Red/Yellow/Green)

```python
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import PatternFill

green = PatternFill(bgColor='C6EFCE')
yellow = PatternFill(bgColor='FFEB9C')
red = PatternFill(bgColor='FFC7CE')

ws.conditional_formatting.add('D2:D100',
    CellIsRule(operator='greaterThanOrEqual', formula=['0.1'], fill=green))
ws.conditional_formatting.add('D2:D100',
    CellIsRule(operator='between', formula=['0', '0.1'], fill=yellow))
ws.conditional_formatting.add('D2:D100',
    CellIsRule(operator='lessThan', formula=['0'], fill=red))
```

### Data Bars

```python
from openpyxl.formatting.rule import DataBarRule

ws.conditional_formatting.add('C2:C100',
    DataBarRule(start_type='min', end_type='max', color='638EC6'))
```

### Color Scale

```python
from openpyxl.formatting.rule import ColorScaleRule

ws.conditional_formatting.add('E2:E100',
    ColorScaleRule(start_type='min', start_color='F8696B',
                   mid_type='percentile', mid_value=50, mid_color='FFEB84',
                   end_type='max', end_color='63BE7B'))
```

---

## Large Dataset Handling {#large-data}

### When to Use Which Approach

```
< 10,000 rows  → openpyxl (full features, formulas, formatting)
10K - 100K rows → pandas write (pd.to_excel) + openpyxl for formatting
100K+ rows      → Write-only mode or CSV output
```

### Write-Only Mode (Memory Efficient)

```python
wb = Workbook(write_only=True)
ws = wb.create_sheet()

for row_data in large_dataset:
    ws.append(row_data)

wb.save('large_output.xlsx')
```

### Pre-compute in Python, Load Static

For 100K+ rows with complex calculations:
1. Do all calculations in pandas/numpy
2. Write results as static values to Excel
3. Add summary formulas only on aggregated data
4. Note in documentation: "Detail rows are static values; summary rows contain live formulas"

### Chunked Processing

```python
chunk_size = 10000
for i, chunk in enumerate(pd.read_csv('huge.csv', chunksize=chunk_size)):
    if i == 0:
        chunk.to_excel('output.xlsx', index=False, sheet_name='Data')
    else:
        # Append rows using openpyxl
        pass
```

---

## Common Templates {#templates}

### Three-Statement Financial Model Skeleton

```
Sheets: Assumptions, Income Statement, Balance Sheet, Cash Flow, Valuation
Assumptions: Growth rates, margins, tax rate, CapEx, D&A, working capital days
Income Statement: Revenue → COGS → Gross Profit → OpEx → EBITDA → D&A → EBIT → Interest → Tax → Net Income
Balance Sheet: Current Assets, PP&E, Intangibles | Current Liabilities, Long-term Debt, Equity
Cash Flow: Net Income → + D&A → ± WC changes → Operating CF | CapEx → Investing CF | Debt → Financing CF
Linkages: Net Income flows to CF; Cash flows to BS; D&A consistent across all three
```

### Dashboard Template

```
Sheets: Data, Calculations, Dashboard
Data: Raw data with headers in row 1
Calculations: SUMIFS, COUNTIFS, AVERAGEIFS aggregating Data
Dashboard: KPI cards (big numbers), trend charts, comparison charts
Layout: Title bar → KPI row → Charts row → Detail table
```

### Budget vs Actual Template

```
Sheets: Assumptions, Budget, Actual, Variance
Columns: Line Item | Budget | Actual | Variance ($) | Variance (%) | Commentary
Variance: =Actual - Budget (favorable/unfavorable logic depends on line type)
Conditional formatting: Red for unfavorable > 10%, Green for favorable
```

### Engineering Calculation Template

```
Sheets: Inputs, Calculations, Results, Verification
Inputs: All parameters with units in headers, data validation on ranges
Calculations: Step-by-step formulas with intermediate results visible
Results: Summary table with final values, units, and pass/fail status
Verification: Python cross-check results, unit consistency log
```
