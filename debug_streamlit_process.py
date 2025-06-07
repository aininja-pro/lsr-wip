#!/usr/bin/env python3

import sys
import os
sys.path.append('/Users/richardrierson/Desktop/Projects/WIP/src')

from data_processing.excel_integration_v2 import find_or_create_monthly_tab, load_wip_workbook, find_section_markers

def simulate_streamlit_upload():
    """Simulate exactly what the Streamlit app does"""
    
    # Find the WIP Report file
    wip_files = []
    test_data_dir = 'test_data'
    if os.path.exists(test_data_dir):
        for file in os.listdir(test_data_dir):
            if 'WIP Report' in file and file.endswith(('.xlsx', '.xlsm')) and not file.startswith('~$'):
                wip_files.append(os.path.join(test_data_dir, file))

    if not wip_files:
        print('No WIP Report files found')
        return
        
    file_path = wip_files[0]
    print(f'Simulating Streamlit process with: {file_path}')
    
    # Step 1: Read file as bytes (like st.file_uploader does)
    with open(file_path, 'rb') as f:
        file_bytes = f.read()
    
    # Step 2: Write to temp file (like Streamlit app does)
    temp_path = "temp_master.xlsx"
    with open(temp_path, "wb") as f:
        f.write(file_bytes)
    print(f'Created temporary file: {temp_path}')
    
    # Step 3: Load workbook (like display_excel_preview does)
    try:
        wb = load_wip_workbook(temp_path)
        print(f'Loaded workbook with sheets: {wb.sheetnames}')
        
        # Step 4: Find/create monthly tab (this is where the issue might be)
        month_year = "Apr 25"  # This is what would be passed from the UI
        ws = find_or_create_monthly_tab(wb, month_year)
        print(f'Found/created tab: {ws.title}')
        
        # Step 5: Find sections (this is where the warnings come from)
        section_patterns = ["5040", "5030"]
        section_markers = find_section_markers(ws, section_patterns)
        print(f'Section markers result: {section_markers}')
        
        # Step 6: Check specific cells manually
        print(f'\nManual cell check:')
        print(f'Cell at Row 2, Col 2: "{ws.cell(2, 2).value}"')
        print(f'Cell at Row 69, Col 2: "{ws.cell(69, 2).value}"')
        
        # Step 7: Debug the search process
        print(f'\nDebug search process:')
        for row in range(1, 5):  # Check first few rows
            for col in range(1, 5):  # Check first few columns
                cell = ws.cell(row=row, column=col)
                if cell.value:
                    print(f'Row {row}, Col {col}: "{cell.value}"')
    
    except Exception as e:
        print(f'Error: {str(e)}')
        import traceback
        traceback.print_exc()
    
    finally:
        # Clean up
        if os.path.exists(temp_path):
            os.remove(temp_path)
            print(f'Cleaned up: {temp_path}')

if __name__ == "__main__":
    simulate_streamlit_upload() 