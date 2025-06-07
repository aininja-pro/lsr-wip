#!/usr/bin/env python3
"""
Test script to check section detection robustness across multiple monthly tabs.
This will reveal how format changes over time and test our detection logic.
"""

import pandas as pd
from pathlib import Path
import sys
import os

# Add src to path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from data_processing.excel_integration_v2 import (
    load_wip_workbook, 
    find_section_markers
)

def count_jobs_in_section(worksheet, start_row, column=1, max_rows=50):
    """
    Count jobs in a section starting from given row.
    
    Args:
        worksheet: The worksheet to scan
        start_row: Row to start counting from
        column: Column to check for job data (default: 1 = Column A)
        max_rows: Maximum rows to scan
        
    Returns:
        tuple: (job_count, jobs_list)
    """
    jobs = []
    consecutive_empty = 0
    
    for row_offset in range(max_rows):
        row_num = start_row + row_offset
        cell = worksheet.cell(row=row_num, column=column)
        
        if cell.value and str(cell.value).strip():
            job_value = str(cell.value).strip()
            # Skip headers and totals
            if not any(skip_word in job_value.lower() for skip_word in 
                      ['job#', 'job description', 'total', 'subtotal', 'sum', '']):
                jobs.append(job_value)
                consecutive_empty = 0
            else:
                consecutive_empty += 1
        else:
            consecutive_empty += 1
        
        # Stop if we hit 3 consecutive empty/header rows
        if consecutive_empty >= 3:
            break
    
    return len(jobs), jobs

def analyze_worksheet_structure(worksheet, sheet_name):
    """Analyze the structure of a single worksheet."""
    print(f"\n📋 Analyzing {sheet_name}...")
    
    # Find section markers
    markers = find_section_markers(worksheet, ['5040', '5030'])
    
    analysis = {
        'sheet_name': sheet_name,
        'sections_found': {},
        'job_counts': {},
        'sample_jobs': {}
    }
    
    # Check 5040 section
    if markers['5040']:
        start_row, start_col = markers['5040']
        print(f"   ✅ 5040 section found at row {start_row}, col {start_col}")
        
        # Count jobs starting from row after header
        job_count, job_list = count_jobs_in_section(worksheet, start_row + 1)
        analysis['sections_found']['5040'] = (start_row, start_col)
        analysis['job_counts']['5040'] = job_count
        analysis['sample_jobs']['5040'] = job_list[:3]  # First 3 jobs as sample
        
        print(f"      Jobs found: {job_count}")
        if job_list:
            print(f"      Sample jobs: {job_list[:3]}")
    else:
        print(f"   ❌ 5040 section not found")
        analysis['sections_found']['5040'] = None
        analysis['job_counts']['5040'] = 0
    
    # Check 5030 section
    if markers['5030']:
        start_row, start_col = markers['5030']
        print(f"   ✅ 5030 section found at row {start_row}, col {start_col}")
        
        # Count jobs starting from row after header
        job_count, job_list = count_jobs_in_section(worksheet, start_row + 1)
        analysis['sections_found']['5030'] = (start_row, start_col)
        analysis['job_counts']['5030'] = job_count
        analysis['sample_jobs']['5030'] = job_list[:3]  # First 3 jobs as sample
        
        print(f"      Jobs found: {job_count}")
        if job_list:
            print(f"      Sample jobs: {job_list[:3]}")
    else:
        print(f"   ❌ 5030 section not found")
        analysis['sections_found']['5030'] = None
        analysis['job_counts']['5030'] = 0
    
    return analysis

def test_multiple_tabs_robustness():
    """Test section detection across multiple monthly tabs."""
    print("🧪 Testing Section Detection Robustness Across Multiple Tabs")
    print("=" * 70)
    
    # Load the Master WIP Report
    master_report_file = Path("test_data") / "Master WIP Report.xlsx"
    
    if not master_report_file.exists():
        print(f"❌ Master WIP Report not found: {master_report_file}")
        return False
    
    try:
        # Load workbook
        workbook = load_wip_workbook(str(master_report_file))
        print(f"✅ Loaded workbook with {len(workbook.sheetnames)} sheets")
        
        # Target monthly tabs to test (in chronological order)
        target_tabs = [
            'Nov 23', 'Dec 23', 'Jan 24', 'Feb 24', 'March 24', 'April 24', 
            'May 24', 'June 24', 'July 24', 'August 24', 'Sept 24', 'Oct 24',
            'Nov 24', 'Dec 24', 'Jan 25', 'Feb 25', 'March 25', 'April 25', 'May 25'
        ]
        
        # Find which tabs actually exist
        existing_tabs = [tab for tab in target_tabs if tab in workbook.sheetnames]
        print(f"📊 Found {len(existing_tabs)} monthly tabs to test: {existing_tabs}")
        
        # Analyze each tab
        results = []
        
        for tab_name in existing_tabs:
            worksheet = workbook[tab_name]
            analysis = analyze_worksheet_structure(worksheet, tab_name)
            results.append(analysis)
        
        # Summary analysis
        print(f"\n📈 SUMMARY ANALYSIS")
        print("=" * 50)
        
        print(f"Tabs tested: {len(results)}")
        
        # Section detection success rate
        tabs_with_5040 = sum(1 for r in results if r['sections_found']['5040'])
        tabs_with_5030 = sum(1 for r in results if r['sections_found']['5030'])
        
        print(f"5040 sections detected: {tabs_with_5040}/{len(results)} ({tabs_with_5040/len(results)*100:.1f}%)")
        print(f"5030 sections detected: {tabs_with_5030}/{len(results)} ({tabs_with_5030/len(results)*100:.1f}%)")
        
        # Job count analysis
        print(f"\n📊 Job Count Analysis:")
        print("Tab Name".ljust(15) + "5040 Jobs".ljust(12) + "5030 Jobs".ljust(12) + "Notes")
        print("-" * 50)
        
        for result in results:
            tab_name = result['sheet_name']
            count_5040 = result['job_counts']['5040']
            count_5030 = result['job_counts']['5030']
            
            notes = []
            if result['sections_found']['5040'] is None:
                notes.append("No 5040")
            if result['sections_found']['5030'] is None:
                notes.append("No 5030")
            
            notes_str = ", ".join(notes) if notes else "OK"
            
            print(f"{tab_name:<15}{count_5040:<12}{count_5030:<12}{notes_str}")
        
        # Section position analysis
        print(f"\n📍 Section Position Analysis:")
        print("Tab Name".ljust(15) + "5040 Position".ljust(18) + "5030 Position")
        print("-" * 50)
        
        for result in results:
            tab_name = result['sheet_name']
            pos_5040 = result['sections_found']['5040']
            pos_5030 = result['sections_found']['5030']
            
            pos_5040_str = f"Row {pos_5040[0]}, Col {pos_5040[1]}" if pos_5040 else "Not found"
            pos_5030_str = f"Row {pos_5030[0]}, Col {pos_5030[1]}" if pos_5030 else "Not found"
            
            print(f"{tab_name:<15}{pos_5040_str:<18}{pos_5030_str}")
        
        # Consistency check
        print(f"\n🔍 Format Consistency Check:")
        
        # Check if 5040 sections are in consistent positions
        valid_5040_positions = [r['sections_found']['5040'] for r in results if r['sections_found']['5040']]
        valid_5030_positions = [r['sections_found']['5030'] for r in results if r['sections_found']['5030']]
        
        if valid_5040_positions:
            unique_5040_rows = set(pos[0] for pos in valid_5040_positions)
            print(f"5040 section found at rows: {sorted(unique_5040_rows)}")
            if len(unique_5040_rows) == 1:
                print("✅ 5040 section position is consistent")
            else:
                print("⚠️  5040 section position varies between tabs")
        
        if valid_5030_positions:
            unique_5030_rows = set(pos[0] for pos in valid_5030_positions)
            print(f"5030 section found at rows: {sorted(unique_5030_rows)}")
            if len(unique_5030_rows) == 1:
                print("✅ 5030 section position is consistent")
            else:
                print("⚠️  5030 section position varies between tabs")
        
        # Detection robustness assessment
        detection_success_rate = (tabs_with_5040 + tabs_with_5030) / (len(results) * 2) * 100
        
        print(f"\n🎯 ROBUSTNESS ASSESSMENT:")
        print(f"Overall detection success rate: {detection_success_rate:.1f}%")
        
        if detection_success_rate >= 90:
            print("✅ EXCELLENT: Detection logic is very robust")
        elif detection_success_rate >= 75:
            print("✅ GOOD: Detection logic is fairly robust")
        elif detection_success_rate >= 50:
            print("⚠️  FAIR: Detection logic needs improvement")
        else:
            print("❌ POOR: Detection logic needs significant work")
        
        # Recommendations
        print(f"\n💡 RECOMMENDATIONS:")
        if detection_success_rate < 100:
            print("• Add preview step in Streamlit interface to confirm detected sections")
            print("• Allow manual override for section positions")
            print("• Consider adding more section header patterns to search for")
        
        if len(set(pos[0] for pos in valid_5040_positions)) > 1:
            print("• 5040 section position varies - ensure dynamic detection is working")
        
        if len(set(pos[0] for pos in valid_5030_positions)) > 1:
            print("• 5030 section position varies - ensure dynamic detection is working")
        
        return True
        
    except Exception as e:
        print(f"❌ ERROR during testing: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_multiple_tabs_robustness()
    if success:
        print("\n🚀 Multi-tab robustness test completed!")
    else:
        print("\n⚠️  Issues found during testing") 