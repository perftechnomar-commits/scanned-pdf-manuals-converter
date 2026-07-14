# Spare Parts OCR Import Builder

This Streamlit project converts scanned spare-parts manuals into reviewed rows and then writes the approved records into the supplied Benefit Excel import template.

## Included files

- `app.py` — Streamlit interface.
- `benefit_tools.py` — OCR, structured extraction, validation, and Excel-generation functions.
- `Spare parts template last version.xlsx` — the original template supplied for this project.
- `requirements.txt` — Python dependencies.
- `.streamlit/secrets.toml.example` — API-key example.

## Run locally

Use Python 3.12, which matches the current `py-mistral-helper` project requirement.

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
# source .venv/bin/activate

pip install -r requirements.txt
```

Copy the example secrets file and add your key:

```text
.streamlit/secrets.toml
```

```toml
MISTRAL_API_KEY = "your-key"
```

Start the app:

```bash
streamlit run app.py
```

## Workflow

1. Enter the main machinery details in the **Machinery** tab.
2. Keep **PDF** selected, upload a scanned manual, and optionally choose a page range such as `1-25`.
3. Run OCR. The app can process a large manual in page chunks.
4. The recommended mode sends the OCR markdown to a Mistral text model using JSON mode to produce structured spare-parts rows. A local markdown-table parser is available as a fallback.
5. Correct identifiers, descriptions, machinery names, units, and quantities in the **Review spare parts** tab.
6. Rows must be both `INCLUDE = True` and `READY = True` before export.
7. Create and download the workbook from the **Export** tab.

## Mapping

### Sheet `1.Machineries|Sub|Units`, starting at row 5

`CODE`, `NAME`, `MAKER`, `MODEL`, `TYPE`, `INSTR.BOOK`, `SPECIFICATIONS`, `MCH_TP(M/S/U)`

### Sheet `2.Spare Parts`, starting at row 4

`MACHINERY`, `PART NO`, `DESCRIPTION`, `CODE`, `ITEM NO`, `UNIT`, `QNT`

The first column of the spare-parts sheet has a blank visible header in the supplied template, but it is the machinery-selection column.

## Important controls

- The original template formatting and sheets are retained; values are written into the pre-existing data areas.
- Part numbers, codes, and item numbers are written as text to preserve leading zeros.
- Rows with a missing machinery, missing description, or neither a part number nor item number are blocked.
- A spare-part machinery value must exactly match a `NAME` on the machinery sheet.
- Possible duplicates are blocked by default.
- Low-confidence OCR rows are not automatically blocked, but they are clearly flagged for manual verification.
- Use the separate audit workbook to retain source page numbers and OCR text; those audit fields are not added to import workbook.

Always test a small import batch before processing a complete manual.
