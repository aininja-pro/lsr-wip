# Blueprint Update: Lien Letter Generator — Phase 2 Changes

## Summary

Update the Lien Letter Generator to handle the new spreadsheet format, generate three letter variants per invoice (Customer, Manager, Owner), add invoice number to letters, and produce Avery 5160 mailing labels. This is a modification of the existing Phase 1 implementation, not a rebuild.

**Read the original `blueprint.md` first.** This document only describes what changes. Anything not mentioned here stays the same.

---

## Change 1: New Spreadsheet Format

The spreadsheet format has changed. **Replace the entire parsing logic in `letter_processing.py`.**

### New Column Layout (by position)

Row 1 is the header row. There are no metadata rows to skip.

| Position | Column Letter | Header Name | Canonical Name | Used For |
|----------|--------------|-------------|----------------|----------|
| 0 | A | Parent Account | `manager_name` | Manager letter |
| 1 | B | Parent Account Address | `manager_address` | Manager letter |
| 2 | C | Suite | `manager_suite` | Manager letter |
| 3 | D | City | `manager_city` | Manager letter |
| 4 | E | State | `manager_state` | Manager letter |
| 5 | F | Zip Code | `manager_zip` | Manager letter |
| 6 | G | Customer Name | `customer_name` | Customer letter + RE block on all letters |
| 7 | H | Service Location Address 1 | `service_address` | Customer letter + RE block on all letters |
| 8 | I | Service Location City | `service_city` | Customer letter + RE block on all letters |
| 9 | J | State | `service_state` | Customer letter + RE block on all letters |
| 10 | K | Zip Code | `service_zip` | Customer letter + RE block on all letters |
| 11 | L | Invoice# | `invoice_number` | Invoice bullet in all letters |
| 12 | M | Invoice Date | `invoice_date` | Not used in letter (keep in data) |
| 13 | N | Invoice Total | `invoice_total` | Invoice bullet in all letters |
| 14 | O | Owner | `owner_name` | Owner letter |
| 15 | P | Owner Address | `owner_address` | Owner letter |
| 16 | Q | Suite | `owner_suite` | Owner letter |
| 17 | R | City | `owner_city` | Owner letter |
| 18 | S | State | `owner_state` | Owner letter |
| 19 | T | Zip | `owner_zip` | Owner letter |

### Parsing Changes

- **No metadata rows.** Row 1 is the header. Read data starting from row 2.
- **Still use fuzzy column matching** via `column_mapping.py` for the header names, but also support positional fallback. If fuzzy matching fails (e.g., duplicate "State" and "Zip Code" headers), fall back to column position as defined above.
- **Suite fields are new.** These did not exist in Phase 1. Add them to column mappings.
- **Invoice # is now in Column L.** Map it. This fixes the "NO_INV" issue from Phase 1.

### Owner Field Logic

For each row, evaluate the Owner column (position 14) to determine if an owner letter is needed:

1. **Owner is blank/None** → No owner letter. This is normal.
2. **Owner matches "SAME AS MANAGEMENT"** → No owner letter. Match case-insensitively and handle variations:
   - "SAME AS MANAGEMENT"
   - "Same as Management"
   - "Same as Mgmt"
   - "Same as mgmt"
   - "SAME"
   - "Same"
   - Strategy: lowercase the value, strip whitespace. If it starts with "same", treat as no owner letter.
3. **Owner has a name AND a complete address** (owner_address + owner_city + owner_state + owner_zip all present) → Generate owner letter.
4. **Owner has partial data** (name but missing address fields, or address fields but missing name) → **Do NOT generate owner letter. Flag as warning.** Warning message: "Row {N} ({customer_name}): Owner data is incomplete — has {what's present} but missing {what's missing}. No owner letter generated. Please verify."

### Updated Return Value

```
parse_invoice_spreadsheet(uploaded_file) -> tuple[pd.DataFrame, list[str]]
```

Same signature, but the DataFrame now includes all 20 columns mapped to canonical names, plus a computed column:

- `owner_letter_needed` — boolean. True only if owner data is complete (case 3 above).

---

## Change 2: Three Letter Variants Per Invoice

### Letter Generation Logic

For each row in the parsed DataFrame, generate up to 3 letters:

**Letter 1: Customer Letter (always generated)**
- Address block: `customer_name`, `service_address`, `service_city`, `service_state`, `service_zip`
- No suite field for customer (not in spreadsheet)
- RE block: Property = `customer_name`, Location = `service_address, service_city, service_state service_zip`

**Letter 2: Manager Letter (always generated)**
- Address block: `manager_name`, `manager_address`, `manager_suite`, `manager_city`, `manager_state`, `manager_zip`
- RE block: Same as customer letter (property name and service location)

**Letter 3: Owner Letter (only if `owner_letter_needed` is True)**
- Address block: `owner_name`, `owner_address`, `owner_suite`, `owner_city`, `owner_state`, `owner_zip`
- RE block: Same as customer letter (property name and service location)

### Address Block Format (all letter types)

```
{name}
{address}, {suite}          ← suite on same line as address, comma-separated
{city}, {state} {zip}
```

If suite is blank/None, omit it and the comma:

```
{name}
{address}
{city}, {state} {zip}
```

### RE Block Format (same on all 3 variants)

```
RE:
Property: {customer_name}
Location: {service_address}, {service_city}, {service_state} {service_zip}
```

### Updated Function Signature

Update `generate_letter()` in `letter_template.py`:

```
generate_letter(
    row: dict,
    recipient_type: str,      # "Customer", "Manager", or "Owner"
    letter_date: date,
    deadline_date: date,
    logo_path: str
) -> BytesIO
```

The function selects the correct address fields based on `recipient_type`:

| recipient_type | Name field | Address field | Suite field | City field | State field | Zip field |
|---|---|---|---|---|---|---|
| "Customer" | `customer_name` | `service_address` | (none) | `service_city` | `service_state` | `service_zip` |
| "Manager" | `manager_name` | `manager_address` | `manager_suite` | `manager_city` | `manager_state` | `manager_zip` |
| "Owner" | `owner_name` | `owner_address` | `owner_suite` | `owner_city` | `owner_state` | `owner_zip` |

Everything else in the letter (salutation, body paragraphs, invoice bullet, deadline, signature) is identical across all 3 variants.

### File Naming

```
Lien_Notice_{customer_name}_{invoice_number}_{recipient_type}.docx
Lien_Notice_{customer_name}_{invoice_number}_{recipient_type}.pdf
```

Examples:
- `Lien_Notice_Lex_123896_Customer.docx`
- `Lien_Notice_Lex_123896_Manager.docx`
- `Lien_Notice_Lex_123896_Owner.docx`

Same sanitization rules as Phase 1 for customer_name (replace spaces with underscores, strip special characters).

---

## Change 3: ZIP Structure

Replace the Phase 1 ZIP structure (flat `docx/` and `pdf/` folders) with invoice-grouped folders:

```
letters.zip
├── 123896/
│   ├── Lien_Notice_Lex_123896_Customer.docx
│   ├── Lien_Notice_Lex_123896_Customer.pdf
│   ├── Lien_Notice_Lex_123896_Manager.docx
│   ├── Lien_Notice_Lex_123896_Manager.pdf
│   ├── Lien_Notice_Lex_123896_Owner.docx
│   └── Lien_Notice_Lex_123896_Owner.pdf
├── 124067/
│   ├── Lien_Notice_Prairie_Commons_124067_Customer.docx
│   ├── Lien_Notice_Prairie_Commons_124067_Customer.pdf
│   ├── Lien_Notice_Prairie_Commons_124067_Manager.docx
│   └── Lien_Notice_Prairie_Commons_124067_Manager.pdf
├── labels.docx
└── labels.pdf
```

- Folder name = invoice number
- 2 files per recipient (docx + pdf)
- 2-3 recipients per invoice (Customer + Manager always, Owner when applicable)
- Labels at the ZIP root (one document for entire batch)

---

## Change 4: Avery 5160 Mailing Labels

### New Module: `src/letter_generation/label_generator.py`

Create a new module with one public function:

```
generate_labels(rows: pd.DataFrame) -> BytesIO
```

**Returns:** A BytesIO object containing a .docx file formatted as Avery 5160 labels.

### Avery 5160 Specifications

- Page: US Letter (8.5" × 11")
- Labels per page: 30 (3 columns × 10 rows)
- Label size: 2.625" wide × 1" tall
- Top margin: 0.5"
- Side margin: 0.1875" (3/16")
- Horizontal gap between labels: 0.125" (1/8")
- Vertical gap: 0 (labels are contiguous vertically)

### Implementation

Use `python-docx` to create a Word document with a table that matches the Avery 5160 grid. Each page is a 10-row × 3-column table with cell dimensions matching the label specs. Remove all table borders (labels should print without visible grid lines).

### Label Content Per Cell

Each label contains:

```
{name}
{address}, {suite}
{city}, {state} {zip}
```

If suite is blank, omit it and the comma. Same format as the letter address block.

Font: Arial or Calibri, 10pt. Left-aligned with small cell padding.

### Label Order and Grouping

For each row in the spreadsheet, generate labels in this order:

1. **3 copies** of Customer label (customer_name, service_address, service_city, service_state, service_zip)
2. **3 copies** of Manager label (manager_name, manager_address + manager_suite, manager_city, manager_state, manager_zip)
3. **3 copies** of Owner label (only if owner_letter_needed is True) (owner_name, owner_address + owner_suite, owner_city, owner_state, owner_zip)

So each invoice row produces 6 labels (no owner) or 9 labels (with owner).

Labels fill left-to-right, top-to-bottom across the 3-column grid. When a page fills (30 labels), start a new page.

### Label Document Output

- Generate as `.docx`
- Also convert to `.pdf` using the same LibreOffice conversion as letters
- Both go in the ZIP root: `labels.docx` and `labels.pdf`

---

## Change 5: Updated Summary Table

Update the summary table displayed in the UI after generation:

| Column | Source |
|---|---|
| # | Row number (1-based) |
| Invoice # | `invoice_number` |
| Property | `customer_name` |
| Manager | `manager_name` |
| Owner | `owner_name` or "—" if blank or "Same as Mgmt" |
| Amount | `invoice_total` (formatted as currency) |
| Letters | Count: "3" or "2" depending on owner |
| Status | "Generated" or "Warning: {reason}" |

### Updated End-of-Run Summary

```
✓ Generated {N_letters} letters for {N_invoices} invoices ({N_success} successful, {N_errors} errors)
  - Customer letters: {N}
  - Manager letters: {N}
  - Owner letters: {N}
  - Labels: {N_labels} across {N_pages} pages
Total amount: ${total:,.2f}
Payment deadline: {deadline_date}
```

---

## Change 6: Column Mapping Updates

Add these new mappings to `column_mapping.py`:

| Canonical Name | Possible Variations |
|---|---|
| `manager_name` | Parent Account, Parent Account Name, Management Company, Manager |
| `manager_address` | Parent Account Address, Manager Address, Mgmt Address |
| `manager_suite` | Suite (position 2 — first Suite column) |
| `manager_city` | City (position 3 — first City column) |
| `manager_state` | State (position 4 — first State column) |
| `manager_zip` | Zip Code (position 5 — first Zip column) |
| `owner_name` | Owner, Property Owner, Owner Name |
| `owner_address` | Owner Address |
| `owner_suite` | Suite (position 16 — third Suite column) |
| `owner_city` | City (position 17 — third City column) |
| `owner_state` | State (position 18 — third State column) |
| `owner_zip` | Zip, Owner Zip, Zip Code (position 19 — third Zip column) |

**Critical:** The spreadsheet has duplicate header names (two "State" columns, two "Zip Code" columns, two "Suite" columns). Fuzzy matching alone will not resolve these. **Use positional index as the primary key for ambiguous columns.** The column position table in Change 1 is authoritative.

---

## Implementation Order

1. **Update column mappings** — Add new canonical names, add positional fallback logic for duplicate headers
2. **Update spreadsheet parser** — New format, no metadata rows, owner logic, suite fields
3. **Update letter template** — Add `recipient_type` parameter, address block switching, suite handling, ensure invoice # is populated
4. **Build label generator** — New module `src/letter_generation/label_generator.py`
5. **Update ZIP builder** — Invoice-grouped folder structure, include labels at root
6. **Update UI** — New summary table columns, updated end-of-run stats
7. **Integration test** — Use `Template_for_30_day_Notices.xlsx` as test input. Verify:
   - Invoice 123896 (Lex) produces 3 folders of letters (has distinct owner)
   - Invoice 123801 (Dylan Apartments) produces 2 folders (owner is blank)
   - Invoice 123872 (Residence at North Dallas) produces 2 folders (owner is "SAME AS MANAGEMENT")
   - Row 16 (Parkwyn, invoice 123946) flags warning for missing owner name with address present
   - Row 17 (Mission Fairways, invoice 123745) flags warning for owner name with no address
   - Invoice numbers appear correctly in all letters
   - Labels are correctly formatted as Avery 5160
   - Label count matches: 3 per recipient per invoice

---

## Files Modified

| File | Action |
|---|---|
| `src/data_processing/column_mapping.py` | MODIFY — add new mappings + positional fallback |
| `src/data_processing/letter_processing.py` | MODIFY — new spreadsheet format, owner logic, suite fields |
| `src/letter_generation/letter_template.py` | MODIFY — recipient_type parameter, address switching, suite, invoice # fix |
| `src/letter_generation/label_generator.py` | NEW — Avery 5160 label document generator |
| `src/ui/app_safe_report.py` | MODIFY — updated summary table, ZIP structure, label inclusion |

**Do not modify any other files.**
