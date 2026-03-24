# Requirements: LSR_Pay Phase 2 — Lien Letter Generator + UI Upgrade

## Problem

LSR Multifamily's AR specialist manually creates pre-lien notice letters for overdue invoices. She works from a spreadsheet export of 90+ day delinquent accounts, opens a Word template, and manually fills in the property name, address, invoice number, and amount for each letter — one at a time. At 80-100+ invoices per batch, this is hours of repetitive copy-paste work with high error risk (wrong amount, wrong address, wrong invoice number on a legal notice).

Separately, the existing LSR_Pay WIP report tool works but looks like a default Streamlit prototype. It doesn't reflect the professional quality expected of a tool used in a business context.

## Cost of Status Quo

- **Letter generation:** 3-5 hours per batch of manual copy-paste into Word templates. Error-prone on a document with legal implications (mechanic's lien notices).
- **UI quality:** The app works but looks unfinished. Undermines confidence in the tool and in 120x.ai's deliverables.

## Who This Is For

- **Primary user:** Stephanie Campbell, AR Specialist, LSR Multifamily
- **Stakeholders:** LSR Multifamily leadership
- **Future users:** Other AR staff if LSR scales

## Proposal

Add a second tool to the existing LSR_Pay Streamlit app: a Lien Letter Generator. The user uploads a spreadsheet of overdue invoices, configures a payment deadline, and the app generates one pre-lien notice letter per invoice row as downloadable Word documents (.docx) and PDFs, bundled in a ZIP file. A summary table lets her review the batch before downloading.

Simultaneously, upgrade the app's visual design using `streamlit-shadcn-ui` components and custom CSS to give the entire app (both WIP Reports and Lien Letters) a clean, professional appearance.

## Success Criteria

- User uploads a spreadsheet and generates 80+ letters in under 2 minutes
- Each letter matches the approved template format with correct property, address, invoice #, and amount
- Letters are downloadable as both .docx (editable) and .pdf
- Summary table displays all letters in the batch for visual QA before download
- App navigation cleanly separates WIP Reports and Lien Letter Generator
- UI looks professional and intentional — not default Streamlit

## Scope

### Included (Phase 1)

**Lien Letter Generator:**
- Upload spreadsheet (.xlsx) of overdue invoices
- Parse spreadsheet data (with fuzzy column matching, consistent with existing app pattern)
- Generate one .docx letter per invoice row using approved template
- Embed LSR Multifamily logo in each letter
- Convert each .docx to PDF
- Display summary table (property name, address, invoice #, amount) for batch QA
- Bundle all .docx and .pdf files into a single ZIP for download
- Configurable payment deadline (days from generation date, default 23)
- Letter date = date of generation

**UI Upgrade:**
- Install and integrate `streamlit-shadcn-ui`
- Tab/navigation component to switch between WIP Reports and Lien Letter Generator
- Card-style layouts for upload areas and settings
- Consistent color scheme, typography, and spacing
- Professional visual hierarchy across both tools
- Custom CSS theming for overall polish

### Included (Phase 2+)

- Owner field placement in letter (pending client direction on where it goes)
- Grouped letters (multiple invoices per property in one letter)
- Email sending integration (SMTP)
- Letter history / audit log
- Editable letter preview within the app

### Not Included

- Direct email sending from the app (Phase 2+)
- Master WIP file modification (existing constraint — safe report pattern only)
- In-app letter text editor
- Any changes to legacy app files

## Dependencies

- Existing LSR_Pay app (`app_safe_report.py`) — this feature extends it
- `streamlit-shadcn-ui` package (new dependency)
- `docx` npm package (for .docx generation via docx-js)
- LibreOffice or equivalent for .docx → PDF conversion
- LSR Multifamily logo file (PNG, to be provided by Ray)
- Updated spreadsheet from Stephanie with Invoice # and Owner columns added

## Inputs

**Spreadsheet (.xlsx) — overdue invoice export:**

| Column | Description | Example |
|--------|-------------|---------|
| Account Number | Internal account code | ELIT-10-01 |
| Parent Account Name | Management company | Greystar |
| Bill To Location Address 1 | Billing address street | 600 East Las Colinas Boulevard |
| Bill To State/Province | Billing address state | TX |
| Bill To Zip/Post Code | Billing address zip | 75039 |
| Customer Name | Property name | Lyndon |
| Service Location Address 1 | Service/property street | 7902 N MacArthur Blvd |
| Service Location City | Service/property city | Irving |
| Service Location State/Province | Service/property state | TX |
| Service Location Zip/Post Code | Service/property zip | 75063 |
| Invoice Date | Date of original invoice | 12/02/2025 |
| Invoice Total | Amount owed | 825 |
| Invoice # | Invoice number (COMING) | 122663 |
| Owner | Property owner (COMING) | TBD |

**Note:** Spreadsheet has 4 metadata rows at top (company name, created date, created by, date range) before the header row. Parser must skip these.

**Logo file:** PNG format, provided separately.

## Outputs

**Per invoice row:**
- One .docx file: pre-lien notice letter matching approved template
- One .pdf file: PDF version of the same letter

**Per batch:**
- ZIP file containing all .docx files in a `/docx` subfolder and all .pdf files in a `/pdf` subfolder
- Summary table displayed in-app before download (property name, service address, invoice #, invoice amount)

**Letter content structure:**
1. LSR Multifamily logo (embedded)
2. Date (generation date)
3. Property name and service location address block
4. RE: block with property name and service location
5. "To Whom It May Concern:" salutation
6. Body paragraph referencing unpaid invoice
7. Invoice bullet: Invoice #{number} - ${amount}
8. Payment deadline (generation date + configurable days)
9. Lien threat paragraph with $250 collection fee
10. Resolution paragraph with contact info
11. Signature block: Stephanie Campbell, AR Specialist, LSR Multifamily, sstorey@lsrusa.com, 972-869-4479

## Constraints

- Must extend `app_safe_report.py` — no new app entry points
- Must follow existing patterns: fuzzy column matching via `column_mapping.py`, safe report (download only, no file modification)
- Docker deployment on Render — any new system dependencies (LibreOffice, Node.js) must be added to Dockerfile
- `streamlit-shadcn-ui` dark mode is unreliable — use light theme
- Do not modify any legacy app files
- Owner field is read from spreadsheet but NOT placed in letters yet (awaiting client direction)

## Open Questions

- **Owner field:** Where does it appear in the letter? Awaiting client input. For now, read it from the spreadsheet but don't use it in the letter template.
- **Logo file:** Ray to provide the PNG. What are the exact dimensions / aspect ratio needed?
- **Invoice # column name:** Assumed to be "Invoice #" — need to confirm with Stephanie's updated spreadsheet.
- **Payment deadline default:** Set at 23 days based on the sample letter. Need to confirm with Stephanie if this is her standard.
- **Signature block:** The sample letter shows Stephanie Campbell but the email is sstorey@lsrusa.com (Stephanie Storey?). Which name/email is correct? Or does this vary?
- **PDF conversion method:** Need to confirm LibreOffice is available in the Docker image, or determine alternative (e.g., `docx2pdf`, WeasyPrint). This affects the Dockerfile.
