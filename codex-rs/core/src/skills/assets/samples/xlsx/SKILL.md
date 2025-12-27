---
name: xlsx
description: Read, write, and manipulate Excel spreadsheets (.xlsx files). Use when working with spreadsheet data, creating reports, extracting data from Excel files, or automating spreadsheet tasks.
---

# Excel Spreadsheet Handling

## Overview

Work with Excel files (.xlsx) using Python's openpyxl library. This skill covers reading data, writing new spreadsheets, and modifying existing files.

## Dependencies

```bash
pip install openpyxl
```

## Common Operations

### Reading an Excel File

```python
from openpyxl import load_workbook

wb = load_workbook('file.xlsx')
ws = wb.active  # or wb['SheetName']

# Read cell value
value = ws['A1'].value

# Iterate rows
for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
    print(row)

# Get all data as list of lists
data = [[cell.value for cell in row] for row in ws.iter_rows()]
```

### Writing a New Excel File

```python
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Data"

# Write cells
ws['A1'] = "Header"
ws.append([1, 2, 3])  # Append row

wb.save('output.xlsx')
```

### Converting to CSV/JSON

Use `scripts/xlsx_tools.py` for quick conversions:

```bash
# To CSV
python scripts/xlsx_tools.py to-csv input.xlsx output.csv

# To JSON
python scripts/xlsx_tools.py to-json input.xlsx output.json

# Get sheet names
python scripts/xlsx_tools.py sheets input.xlsx
```

### Modifying Existing Files

```python
from openpyxl import load_workbook

wb = load_workbook('existing.xlsx')
ws = wb.active

ws['B2'] = "Updated value"
ws.insert_rows(3)  # Insert row at position 3
ws.delete_cols(2)  # Delete column B

wb.save('existing.xlsx')
```

## Formatting

```python
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# Bold header
ws['A1'].font = Font(bold=True, size=14)

# Center alignment
ws['A1'].alignment = Alignment(horizontal='center')

# Background color
ws['A1'].fill = PatternFill(start_color='FFFF00', fill_type='solid')

# Column width
ws.column_dimensions['A'].width = 20
```

## Working with Formulas

```python
ws['C1'] = '=SUM(A1:B1)'
ws['D1'] = '=AVERAGE(A1:A10)'
```

Note: openpyxl preserves formulas but doesn't evaluate them. Use `data_only=True` when loading to get calculated values (requires Excel to have saved the file with cached values).
