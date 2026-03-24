# LSR_Pay — WIP Report Automation Tool

A Streamlit web application that automates monthly WIP (Work-in-Progress) report generation for a construction company. Replaces a ~3 hour manual process of pulling Sage accounting exports, matching job costs against budgets, and updating a master Excel spreadsheet. Reduces to under 5 minutes.

## Tech Stack

- Language: Python 3.12
- UI: Streamlit
- Data Processing: pandas
- Excel: openpyxl (reading), xlsxwriter (writing validation reports)
- Deployment: Docker on Render
- Database: None (fully offline, file-upload based)

## How It Works

User uploads two Sage accounting exports → app processes and merges → outputs a downloadable Excel report.

**Inputs:**
1. GL Inquiry Export (.xlsx) — Contains Account, Job Number, Debit, Credit columns
2. WIP Worksheet Export (.xlsx) — Contains Job Number, Status, Job Description, Contract Amount, Estimated Sub Labor/Material Costs

**Processing:**
- Filter GL rows: Account 5040 (Sub Labor), 5030 (Material), 4020 (Billing)
- Compute: Amount = Debit + Credit; Amount Billed = -Credit (for 4020)
- Aggregate by Job Number + Account Type → pivot to one row per job
- Left-join with WIP Worksheet on Job Number
- Compute variances (Actual - Budget)

**Output:**
- Downloadable Excel workbook with sheets: 5040_Labor_Updates, 5030_Material_Updates, Summary, Instructions
- User manually copies data into their master WIP file (safe report mode — no direct master file modification)

## File Structure

### Production Code (touch these)
```
src/
├── data_processing/
│   ├── column_mapping.py      # Fuzzy column name matching for Sage export variations
│   ├── aggregation.py         # GL data loading, filtering, amount calc, aggregation
│   └── merge_data.py          # WIP worksheet loading, join with GL, variance calc
└── ui/
    └── app_safe_report.py     # PRODUCTION APP — deployed via Docker on Render
```

### Infrastructure
```
Dockerfile                     # Runs app_safe_report.py, WORKDIR /app
requirements.txt               # pandas, openpyxl, streamlit, xlsxwriter
README.md                      # User-facing documentation
WIP_UPDATE_MAPPING.md          # Data flow and column mapping docs (NOTE: 5040/5030 labels are swapped vs code — code is correct)
```

### Legacy — DO NOT MODIFY
These are iteration artifacts from debugging Excel corruption issues. They are not deployed. Do not modify, reference, or extend them.
```
src/ui/app.py                  # Original full-featured version (direct Excel update)
src/ui/app_safe_report_fixed.py # Fixed import paths (not deployed)
src/ui/app_surgical.py         # ZIP/XML surgical Excel approach
src/ui/app_memory_fix.py       # In-memory processing variant
src/ui/app_simple_safe.py      # Ultra-minimal CSV output
src/ui/app_hybrid_claude.py    # Hybrid read/write approach
src/ui/app_fixed_billing.py    # Billing fix variant
src/data_processing/excel_integration.py      # v1 Excel writer
src/data_processing/excel_integration_v2.py   # v2 Excel writer
src/data_processing/excel_surgical.py         # v3 Excel writer (ZIP/XML)
src/debug_sections.py          # Debug scripts
src/debug_sections_clean.py
src/debug_streamlit_path.py
src/test_download_fix.py
src/test_full_pipeline.py
```

## Commands

| Action | Command |
|--------|---------|
| Run locally | `streamlit run src/ui/app_safe_report.py` |
| Run in Docker | `docker build -t lsr-pay . && docker run -p 8501:8501 lsr-pay` |
| Deploy | Push to main (Render auto-deploys from Dockerfile) |

## Conventions

- All data processing logic lives in `src/data_processing/`. No pandas/openpyxl code in UI files.
- Column name matching uses `column_mapping.py` — do not hardcode column names inline.
- The production app uses the "safe report" pattern: generate a new Excel file for download. Never modify the user's uploaded master file.
- Sage exports have inconsistent column names across versions. Always use fuzzy matching.
- Account codes: 5040 = Sub Labor, 5030 = Material, 4020 = Billing. (WIP_UPDATE_MAPPING.md has these swapped — the code is correct.)

## Avoid

- Do not modify any file in the "Legacy" section above.
- Do not add new app variants. Extend `app_safe_report.py` only.
- Do not hardcode column names — use column_mapping.py.
- Do not attempt direct Master WIP file modification — the safe report pattern exists because openpyxl corrupts the master file's formatting.
- Do not add dependencies without checking requirements.txt first.

## Critical Rules

1. Read the relevant source files before making changes.
2. Check in before major changes.
3. Keep changes simple and minimal.
4. The production entry point is `src/ui/app_safe_report.py` — always.

## Known Issues / Tech Debt

- 8 legacy app variants should be archived or deleted (out of scope for now)
- `app_safe_report.py` line 18 hardcodes `sys.path.append('/app/src')` — works in Docker, breaks locally
- Column mapping logic is duplicated in legacy files (only `column_mapping.py` is canonical)
- `merge_data.py` function signature mismatch with some callers (parameter name `fill_missing_with_zero` vs positional `include_closed`)
- No automated tests
- WIP_UPDATE_MAPPING.md contradicts code on 5040/5030 labels

## Environment

- Render deployment: Docker-based, auto-deploys on push to main
- Docker WORKDIR: /app
- Streamlit port: 8501
