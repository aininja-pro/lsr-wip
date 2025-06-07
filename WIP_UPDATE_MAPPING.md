# WIP Report Update Mapping

## Overview
This document explains exactly what gets updated in your WIP Report and where those values come from.

## Data Flow Summary

```
GL Inquiry Export → Aggregate by Job & Account → Split by Section → Update WIP Report
WIP Worksheet → Provide Job Context & Budget Values
```

## Section Updates

### 5040 Section (Material Costs)
**What gets updated:** Material cost values for each job
**Source:** GL Inquiry Export accounts containing "5040"

**Mapping:**
- **Job Numbers (Column A):** From WIP Worksheet (`Job` column)
- **Material Values (Column B):** From GL Inquiry aggregated amounts
  - **Calculation:** Sum of (Debit + Credit) for all GL accounts containing "5040" grouped by Job Number
  - **Account Types:** Material costs, material purchases, material expenses

### 5030 Section (Labor Costs)  
**What gets updated:** Labor cost values for each job
**Source:** GL Inquiry Export accounts containing "5030"

**Mapping:**
- **Job Numbers (Column A):** From WIP Worksheet (`Job` column)
- **Labor Values (Column B):** From GL Inquiry aggregated amounts
  - **Calculation:** Sum of (Debit + Credit) for all GL accounts containing "5030" grouped by Job Number
  - **Account Types:** Labor costs, payroll expenses, subcontractor labor

## Your Real Data Results

From your actual files, here's what we found:

### GL Inquiry Processing
- **Total transactions:** 48 GL entries
- **Unique jobs with activity:** 9 jobs
- **Account filtering:** Successfully found 5040, 5030, and 4020 accounts
- **Sample results:**
  - Job 10-31-223678: Labor: -$1,437.83, Material: $0
  - Job 10-31-224046: Labor: $8,530.89, Material: $15,012.00
  - Job 10-31-224554: Labor: $1,573.41, Material: $0

### WIP Worksheet Integration
- **Total jobs in system:** 48 jobs
- **Column mapping:** Job → Job Number, Status, Description, Budget amounts
- **Merge process:** Left join ensures all WIP jobs are included, GL amounts added where available

## What Gets Preserved (Never Overwritten)

### Formula Cells
- **Any cell containing formulas** (starting with = or having data_type='f')
- **Calculated columns** (totals, percentages, variance calculations)
- **Cell formatting** (colors, borders, fonts, number formats)
- **Column widths and print settings**
- **VBA macros and external links**

### Only Value Cells Updated
- **Job numbers** in column A of each section (if not a formula)
- **Actual amounts** in column B of each section (if not a formula)
- **Clearing stops** at first completely empty row to avoid damaging template structure

## Processing Logic

1. **Load GL Inquiry** → Filter accounts → Group by Job + Account Type → Sum amounts
2. **Load WIP Worksheet** → Get job list and budget amounts → Filter closed jobs (optional)
3. **Merge data** → Combine GL actuals with WIP budget info
4. **Split by section:**
   - 5040 section gets jobs with Material amounts
   - 5030 section gets jobs with Labor amounts
5. **Update Excel:**
   - Create backup with timestamp
   - Find/create monthly tab (e.g., "Jun 25")
   - Locate 5040 and 5030 section markers
   - Clear only value cells (preserve formulas)
   - Write job numbers and amounts
   - Save with VBA preservation

## Monthly Workflow

1. **Export from Sage:** GL Inquiry + WIP Worksheet
2. **Upload to tool:** 3 Excel files (including Master WIP Report)
3. **Select month:** Tool creates/updates appropriate monthly tab
4. **Preview changes:** See old vs new values side-by-side
5. **Apply updates:** Backup created automatically, values updated, formulas preserved
6. **Download results:** Updated WIP Report + validation report with variance flags

## Questions for Verification

Before we build the interface, please confirm:

1. **Column positions:** Are Job Numbers always in column A and amounts in column B for both sections?
2. **Section headers:** Do your monthly tabs have "5040" and "5030" as section headers?
3. **Monthly tab naming:** Should we use "Jun 25" format or do you prefer "June 25"?
4. **Closed jobs:** For quarterly true-ups, should we include all jobs regardless of status?
5. **Formula columns:** Are there specific columns (like C, D, E) that contain formulas we should never touch? 