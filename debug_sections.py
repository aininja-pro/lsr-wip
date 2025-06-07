#!/usr/bin/env python3
import sys
sys.path.append('/app/src')
from data_processing.excel_integration_v2 import find_section_markers, load_wip_workbook
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Load the workbook
print("Loading workbook...")
wb = load_wip_workbook('/app/test_data/Master WIP Report.xlsx')
print(f"Available sheets: {wb.sheetnames}")

# Get the April 25 sheet
ws = wb['April 25']
print(f"Working with sheet: {ws.title}")
print(f"Sheet dimensions: {ws.max_row} rows x {ws.max_column} columns")

# Check specific cells where we know the headers are
print("\nChecking known header locations:")
cell_2_2 = ws.cell(row=2, column=2)
print(f"Row 2, Col 2: '{cell_2_2.value}'")

cell_69_2 = ws.cell(row=69, column=2)
print(f"Row 69, Col 2: '{cell_69_2.value}'")

# Test the section detection
print("\nTesting section detection...")
result = find_section_markers(ws, ['5040', '5030'])
print(f"Section detection result: {result}")

# Manual search for debugging
print("\nManual search for '5040':")
for row in range(1, 20):
    for col in range(1, 10):
        cell = ws.cell(row=row, column=col)
        if cell.value and '5040' in str(cell.value):
            print(f"Found '5040' at row {row}, col {col}: '{cell.value}'")

print("\nManual search for '5030':")
for row in range(60, 80):
    for col in range(1, 10):
        cell = ws.cell(row=row, column=col)
        if cell.value and '5030' in str(cell.value):
            print(f"Found '5030' at row {row}, col {col}: '{cell.value}'") 