#!/usr/bin/env python3

import sys
import os
sys.path.append('/Users/richardrierson/Desktop/Projects/WIP/src')

from data_processing.excel_integration_v2 import find_or_create_monthly_tab, load_wip_workbook, find_section_markers

def debug_wip_file():
    # Look for the WIP Report file
    wip_files = []
    test_data_dir = 'test_data'
    if os.path.exists(test_data_dir):
        for file in os.listdir(test_data_dir):
            if 'WIP Report' in file and file.endswith(('.xlsx', '.xlsm')) and not file.startswith('~$'):
                wip_files.append(os.path.join(test_data_dir, file))

    print(f'Found WIP Report files: {wip_files}')

    if wip_files:
        file_path = wip_files[0]
        print(f'Analyzing: {file_path}')
        
        # Load workbook
        wb = load_wip_workbook(file_path)
        print(f'Sheet names: {wb.sheetnames}')
        
        # Try to find or create Apr 25 tab
        tab_name = 'Apr 25'
        ws = find_or_create_monthly_tab(wb, tab_name)
        
        if ws:
            print(f'Successfully found/created tab: {ws.title}')
            
            # Search for section markers manually
            print('\nSearching for section markers...')
            
            # Look for 5040 section
            for row_idx, row in enumerate(ws.iter_rows(max_row=100), 1):
                for col_idx, cell in enumerate(row, 1):
                    if cell.value and '5040' in str(cell.value):
                        print(f'Found 5040 at Row {row_idx}, Col {col_idx}: {cell.value}')
            
            # Look for 5030 section  
            for row_idx, row in enumerate(ws.iter_rows(max_row=100), 1):
                for col_idx, cell in enumerate(row, 1):
                    if cell.value and '5030' in str(cell.value):
                        print(f'Found 5030 at Row {row_idx}, Col {col_idx}: {cell.value}')
            
            # Use the find_section_markers function with patterns
            section_patterns = {'5040': '5040', '5030': '5030'}
            markers = find_section_markers(ws, section_patterns)
            print(f'\nSection markers found: {markers}')
            
            # Show exactly what's in the cells at those positions
            print(f'\nCell at Row 2, Col 2: "{ws.cell(2, 2).value}"')
            print(f'Cell at Row 69, Col 2: "{ws.cell(69, 2).value}"')
        else:
            print('Could not find or create worksheet')
    else:
        print('No WIP Report files found')

if __name__ == "__main__":
    debug_wip_file() 