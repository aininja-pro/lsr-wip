# Blueprint: LSR_Pay Phase 2 — Lien Letter Generator + UI Upgrade

## Summary

Extend the existing LSR_Pay Streamlit app (`src/ui/app_safe_report.py`) with two changes: (1) a Lien Letter Generator that reads an overdue invoice spreadsheet and produces one pre-lien notice letter per row as .docx and .pdf files bundled in a ZIP, and (2) a UI upgrade using `streamlit-shadcn-ui` components and custom CSS to give the app a professional appearance with tabbed navigation between WIP Reports and Lien Letters.

---

## Dependencies

### New Python Packages (add to requirements.txt)

```
streamlit-shadcn-ui
python-docx
```

### New System Dependencies (add to Dockerfile)

```dockerfile
# LibreOffice for .docx → .pdf conversion
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer && rm -rf /var/lib/apt/lists/*
```

### Existing (already in requirements.txt)

```
streamlit
pandas
openpyxl
xlsxwriter
```

### Static Assets

```
src/assets/lsr_logo.png    # LSR Multifamily logo, 436x120px
```

---

## File Structure — New and Modified Files Only

```
src/
├── assets/
│   └── lsr_logo.png                    # NEW — logo file (Ray places manually)
├── data_processing/
│   ├── column_mapping.py               # MODIFY — add new column mappings for lien letter fields
│   └── letter_processing.py            # NEW — spreadsheet parsing + data extraction for letters
├── letter_generation/
│   ├── letter_template.py              # NEW — .docx letter builder using python-docx
│   └── pdf_converter.py                # NEW — .docx to .pdf conversion via LibreOffice
└── ui/
    ├── app_safe_report.py              # MODIFY — add navigation, integrate lien letter UI, apply theming
    └── styles.py                       # NEW — CSS theme definitions and style injection helper
```

**Do not modify any other files. Do not touch any legacy files.**

---

## Part 1: UI Upgrade

### 1.1 Navigation Structure

Restructure `app_safe_report.py` to use `streamlit-shadcn-ui` tabs as the top-level navigation. The app will have two sections:

```python
import streamlit_shadcn_ui as ui

selected_tab = ui.tabs(
    options=["WIP Reports", "Lien Letters"],
    default_value="WIP Reports",
    key="main_nav"
)

if selected_tab == "WIP Reports":
    render_wip_reports()      # Existing WIP logic, extracted into a function
elif selected_tab == "Lien Letters":
    render_lien_letters()     # New lien letter logic
```

### 1.2 Refactor Existing WIP Code

Move all existing WIP report logic from the body of `app_safe_report.py` into a function called `render_wip_reports()`. This is a pure extraction — no logic changes. The function stays in `app_safe_report.py` (do not create a new file for it).

### 1.3 CSS Theme

Create `src/ui/styles.py` with a function that injects custom CSS via `st.markdown()`. Apply it once at the top of the app, before any content renders.

**Theme direction:**
- Background: clean white / light gray (`#FAFAFA` body, `#FFFFFF` cards)
- Primary color: LSR blue (`#1E3A5F` or close match to the logo's blue)
- Accent: LSR red (`#C41E3A` or close match to the logo's red) — used sparingly for alerts/warnings
- Text: dark gray (`#1A1A1A`) for body, medium gray (`#6B7280`) for secondary text
- Font: system font stack (`-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif`)
- Cards: white background, subtle border (`#E5E7EB`), slight border-radius (`8px`), light shadow
- Spacing: generous padding, clear visual separation between sections

**CSS targets to override in Streamlit:**
- `.stApp` — background color
- `.stTabs` — tab styling (if using native tabs as fallback)
- `h1, h2, h3` — heading fonts and colors
- `.stButton > button` — button styling
- `.stFileUploader` — upload area styling
- `.stDataFrame` — table styling
- `.stDownloadButton` — download button prominence

### 1.4 Streamlit Config

Create or update `.streamlit/config.toml`:

```toml
[theme]
primaryColor = "#1E3A5F"
backgroundColor = "#FAFAFA"
secondaryBackgroundColor = "#FFFFFF"
textColor = "#1A1A1A"
font = "sans serif"
```

### 1.5 Card Layouts

Use `streamlit-shadcn-ui` cards to wrap the upload sections in both tools. Each tool's main area should be inside a card with a title and description.

---

## Part 2: Lien Letter Generator — Data Processing

### 2.1 Column Mapping Updates

Add the following column mappings to `src/data_processing/column_mapping.py`, using the same fuzzy matching pattern already established:

| Canonical Name | Possible Variations |
|---|---|
| `account_number` | Account Number, Acct Number, Account #, Acct # |
| `parent_account_name` | Parent Account Name, Parent Account, Management Company |
| `bill_to_address` | Bill To Location Address 1, Bill To Address, Billing Address |
| `bill_to_state` | Bill To State/Province, Bill To State, Billing State |
| `bill_to_zip` | Bill To Zip/Post Code, Bill To Zip, Billing Zip |
| `customer_name` | Customer Name, Property Name, Property |
| `service_address` | Service Location Address 1, Service Address, Service Location Address |
| `service_city` | Service Location City, Service City |
| `service_state` | Service Location State/Province, Service State |
| `service_zip` | Service Location Zip/Post Code, Service Zip |
| `invoice_date` | Invoice Date, Inv Date |
| `invoice_total` | Invoice Total, Invoice Amount, Amount, Total |
| `invoice_number` | Invoice #, Invoice Number, Inv #, Inv Number |
| `owner` | Owner, Property Owner |

### 2.2 Spreadsheet Parser — `src/data_processing/letter_processing.py`

Create a new module with one public function:

```
parse_invoice_spreadsheet(uploaded_file) -> tuple[pd.DataFrame, list[str]]
```

**Returns:** (DataFrame of parsed invoice rows, list of warning messages)

**Parsing logic:**

1. Read the .xlsx file with `openpyxl` (same as existing GL parser pattern)
2. **Skip metadata rows.** The spreadsheet has up to 4 metadata rows before the header:
   - Row 1: Company name ("LSR Multifamily")
   - Row 2: Created date
   - Row 3: Created by
   - Row 4: Date range
   - Row 5: (possibly blank)
   - Row 6: Column headers
   - Strategy: Scan rows 1-10. The header row is the first row where 3+ cells match known column names from the mapping table above. Read data from the row after that.
3. Apply fuzzy column matching using `column_mapping.py` to normalize column names
4. Drop rows where `invoice_total` is null or zero
5. Strip whitespace from all string fields
6. Normalize `invoice_total` to float
7. Normalize `invoice_number` to string (strip trailing `.0` if read as numeric)
8. Normalize state fields — if full state name, leave as-is. If abbreviation, leave as-is. Don't transform.
9. Normalize zip codes — ensure they're strings, preserve leading zeros (e.g., "01234")
10. **Warnings:** Generate a warning for any row missing a required field (customer_name, service_address, service_city, service_state, service_zip, invoice_total). Include row number in warning. Still include the row in output — let the user decide.
11. **Owner column:** Read it if present. If not present, set to None. Do not warn if missing.

### 2.3 Data Structure

Each row in the returned DataFrame represents one letter to generate. Key columns used by the template:

| Column | Used In Letter | Required |
|---|---|---|
| `customer_name` | Address block, RE block | Yes |
| `service_address` | Address block, RE block | Yes |
| `service_city` | Address block, RE block | Yes |
| `service_state` | Address block, RE block | Yes |
| `service_zip` | Address block, RE block | Yes |
| `invoice_number` | Invoice bullet | Yes |
| `invoice_total` | Invoice bullet | Yes |
| `owner` | Not used yet (Phase 2+) | No |

---

## Part 3: Lien Letter Generator — Document Generation

### 3.1 Letter Template — `src/letter_generation/letter_template.py`

Create a module with one public function:

```
generate_letter(row: dict, letter_date: date, deadline_date: date, logo_path: str) -> BytesIO
```

**Returns:** A BytesIO object containing the .docx file.

**Use `python-docx` (not docx-js).** The existing app is pure Python. Stay in the Python stack. python-docx can handle this letter format without issues.

**Document setup:**
- Page size: US Letter (8.5" × 11")
- Margins: 1" all sides
- Font: Times New Roman or Calibri (match whatever the sample letter uses — inspect the .docx to confirm)
- No headers/footers (the logo goes in the body)

**Letter structure, top to bottom:**

1. **Logo**
   - Insert `lsr_logo.png` from `src/assets/`
   - Width: 4.5 inches (maintain aspect ratio)
   - Left-aligned
   - One blank line after

2. **Date**
   - Format: `February 16, 2026` (full month name, day, 4-digit year)
   - Bold
   - One blank line after

3. **Recipient Address Block**
   - Line 1: `{customer_name}`
   - Line 2: `{service_address}`
   - Line 3: `{service_city}, {service_state} {service_zip}`
   - Regular weight, not bold
   - One blank line after

4. **RE: Block**
   - "RE:" — bold
   - Line 1: `Property: {customer_name}`
   - Line 2: `Location: {service_address}, {service_city}, {service_state} {service_zip}`
   - One blank line after

5. **Salutation**
   - "To Whom It May Concern:" — bold
   - One blank line after

6. **Body Paragraph 1**
   - "We are writing to formally notify you that our company has provided labor and materials for the above-referenced property. According to our records, the following invoice remains unpaid:"
   - Regular weight

7. **Invoice Bullet**
   - Bold: `Invoice #{invoice_number} - ${invoice_total:,.2f}`
   - Use proper bullet list formatting (not a plain dash)
   - Format amount with comma separators and 2 decimal places

8. **Body Paragraph 2**
   - "This notice is provided in accordance with the Texas mechanic's lien laws to protect our legal rights. Payment in full must be received in our office no later than **{deadline_date}**, to avoid further action."
   - Deadline date formatted same as letter date (full month name, day, year)
   - Deadline date should be bold within the paragraph

9. **Body Paragraph 3**
   - "If payment is not received by the deadline stated, we will proceed with filing a lien against the property, and an additional **$250.00** in collection costs will be added to the balance due."
   - "$250.00" should be bold within the paragraph

10. **Body Paragraph 4**
    - "We remain committed to resolving this matter amicably and are available to provide any additional documentation or information you may require. Please contact our office at your earliest convenience to confirm payment arrangements."
    - Regular weight

11. **Signature Block**
    - "Sincerely," followed by two blank lines
    - "Stephanie Campbell" — bold
    - "AR Specialist"
    - "LSR Multifamily"
    - "sstorey@lsrusa.com" — as a clickable mailto link if python-docx supports it, otherwise plain text
    - "972-869-4479"

**File naming convention:** `Lien_Notice_{customer_name}_{invoice_number}.docx`
- Replace spaces with underscores in customer_name
- Strip any characters that aren't alphanumeric, underscore, or hyphen

### 3.2 PDF Conversion — `src/letter_generation/pdf_converter.py`

Create a module with one public function:

```
convert_docx_to_pdf(docx_bytes: BytesIO, output_filename: str) -> BytesIO
```

**Returns:** A BytesIO object containing the PDF.

**Method:** Use LibreOffice in headless mode.

```
libreoffice --headless --convert-to pdf --outdir /tmp/pdf_output /tmp/input.docx
```

**Implementation:**
1. Write the docx BytesIO to a temp file
2. Call LibreOffice via `subprocess.run()`
3. Read the resulting PDF back into BytesIO
4. Clean up temp files
5. If LibreOffice fails, log the error and return None (don't crash the batch — skip the PDF and warn the user)

**Alternative if LibreOffice is too heavy for the Docker image:** Use `docx2pdf` package as a fallback. But try LibreOffice first — it produces the most accurate conversions.

### 3.3 Batch Orchestration

In the UI layer (inside `render_lien_letters()`), after the user clicks "Generate Letters":

1. Parse spreadsheet → get DataFrame + warnings
2. Display warnings if any
3. Show summary table (see 4.3 below)
4. For each row in DataFrame:
   a. Call `generate_letter()` → get .docx BytesIO
   b. Call `convert_docx_to_pdf()` → get .pdf BytesIO
   c. Store both in memory
5. Build ZIP file in memory:
   ```
   letters.zip
   ├── docx/
   │   ├── Lien_Notice_Lyndon_122663.docx
   │   ├── Lien_Notice_Canyon_Grove_122664.docx
   │   └── ...
   └── pdf/
       ├── Lien_Notice_Lyndon_122663.pdf
       ├── Lien_Notice_Canyon_Grove_122664.pdf
       └── ...
   ```
6. Provide ZIP as Streamlit download button

---

## Part 4: Lien Letter Generator — UI

### 4.1 Layout

The Lien Letters tab contains, in order:

1. **Page title:** "Lien Letter Generator" with a subtitle "Generate pre-lien notice letters from an invoice spreadsheet"
2. **Upload card:** File uploader accepting .xlsx only. Wrapped in a shadcn-ui card.
3. **Settings card:** Below the upload card. Contains:
   - "Payment Deadline (days from today)" — number input, default 23, min 7, max 90
   - Display the calculated deadline date as text: "Deadline: {date}"
4. **Generate button:** Prominent button, disabled until a file is uploaded
5. **Warnings area:** If any warnings from parsing, display as a yellow alert/callout
6. **Summary table:** Displayed after generation, before download (see 4.3)
7. **Download button:** "Download All Letters (ZIP)" — large, prominent

### 4.2 Upload and Validation

On file upload:
1. Parse the spreadsheet immediately (call `parse_invoice_spreadsheet()`)
2. If parsing fails entirely (can't find header row, no valid data rows), show error message in red
3. If parsing succeeds with warnings, show warnings but allow proceeding
4. Display a count: "{N} invoices found"
5. Store parsed DataFrame in `st.session_state`

### 4.3 Summary Table

After the user clicks Generate and letters are built, display a table with these columns:

| Column | Source |
|---|---|
| # | Row number (1-based) |
| Property | `customer_name` |
| Address | `{service_address}, {service_city}, {service_state}` |
| Invoice # | `invoice_number` |
| Amount | `invoice_total` (formatted as currency) |
| Status | "Generated" or "Error: {reason}" |

Display using `st.dataframe()` styled with custom CSS for clean appearance.

At the bottom of the table, show a total: "Total: {N} letters, ${total_amount:,.2f}"

### 4.4 State Management

Use `st.session_state` to manage:
- `parsed_df` — the parsed DataFrame from upload
- `parse_warnings` — list of warning strings
- `generated_zip` — the ZIP BytesIO after generation
- `generation_complete` — boolean flag

Clear these when a new file is uploaded.

---

## Part 5: Error Handling

| Scenario | Behavior |
|---|---|
| Uploaded file is not .xlsx | Show error: "Please upload an .xlsx file" |
| Can't find header row | Show error: "Could not find column headers. Expected columns: Customer Name, Invoice Total, etc." |
| No data rows after header | Show error: "No invoice data found in spreadsheet" |
| Row missing required field | Generate warning, include row in output, mark status in summary table |
| Invoice # missing on a row | Use "NO_INV" as placeholder in filename, show warning |
| python-docx fails on a row | Log error, skip that letter, show "Error" status in summary table, continue batch |
| LibreOffice PDF conversion fails | Skip PDF for that letter, include .docx only, show warning |
| LibreOffice not installed | Show error at startup: "PDF conversion requires LibreOffice. DOCX files will still be generated." Generate .docx only. |

**Never crash the batch.** Individual letter failures should be logged and surfaced in the summary table, but the batch continues. The ZIP should contain whatever was successfully generated.

---

## Part 6: Dockerfile Updates

Add to the existing Dockerfile:

```dockerfile
# System dependencies for PDF conversion
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice-writer && \
    rm -rf /var/lib/apt/lists/*
```

**Note:** This will increase the Docker image size significantly (~200-400MB). If this is a problem for Render's build times or storage, fall back to generating .docx only and note PDF conversion as a Phase 2 item. Check Render's build limits before committing to this.

---

## Part 7: Implementation Order

Build in this exact sequence. Test each step before moving to the next.

1. **Column mapping updates** — Add new mappings to `column_mapping.py`. Smallest change, no risk.
2. **Spreadsheet parser** — Build `letter_processing.py`. Test with the Sample.xlsx file. Verify it correctly skips metadata rows, finds headers, and parses all 80+ rows.
3. **Letter template** — Build `letter_template.py`. Generate a single test letter from a hardcoded row dict. Open in Word and compare against Sample_for_Ray.docx visually.
4. **PDF converter** — Build `pdf_converter.py`. Convert the test letter. Verify it opens correctly.
5. **UI — styles and navigation** — Build `styles.py`, add navigation tabs, refactor WIP code into `render_wip_reports()`. Verify WIP Reports still works identically.
6. **UI — Lien Letters tab** — Build `render_lien_letters()` with upload, settings, generation, summary table, and download.
7. **Integration test** — Upload Sample.xlsx through the UI. Verify full pipeline: upload → parse → generate → summary → download ZIP → open letters in Word and as PDF.
8. **Dockerfile update** — Add LibreOffice. Build and test Docker image locally.

---

## Voice and Tone

The generated letters are **formal legal notices**. The tone is:
- Professional and firm, not aggressive
- Clear and specific about amounts and deadlines
- References Texas mechanic's lien law
- Offers resolution ("resolving this matter amicably")
- The letter text is fixed — it comes from the approved template. Do not paraphrase, rearrange, or "improve" the wording. Reproduce it exactly as specified in Section 3.1.

---

## End-of-Run Summary

When letter generation completes, display in the UI:

```
✓ Generated {N} letters ({N_success} successful, {N_errors} errors)
Total amount: ${total:,.2f}
Payment deadline: {deadline_date}
```

If there were errors, also display:
```
⚠ {N_errors} letters had issues — see Status column in the table below
```
