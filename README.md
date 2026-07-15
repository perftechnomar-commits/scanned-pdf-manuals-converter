# Spare Parts OCR Import Builder

A Streamlit application that converts scanned spare-parts manuals into reviewed machinery and spare-parts records, then writes approved data into the bundled Excel import template.

## Main workflow

1. Select one or more vessels.
2. Enter the main machinery code, name, maker, and model.
3. Upload a scanned PDF and run OCR.
4. Review automatically detected sub-machineries and their source-page ranges.
5. Apply the approved sub-machinery assignments to spare-part rows.
6. Review blocked, low-confidence, or unassigned rows using source-page references.
7. Generate the import workbook, audit workbook, and vessel-specific email draft.

## Included files

- `app.py` — Streamlit interface and workflow.
- `tools.py` — OCR, extraction, sub-machinery detection, validation, and Excel generation.
- `vessels.csv` — searchable vessel master list.
- `Spare parts template last version.xlsx` — bundled import template.
- `requirements.txt` — Python dependencies.

## Local setup

Use Python 3.12.

```bash
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS/Linux
# source .venv/bin/activate

pip install -r requirements.txt
```

Create `.streamlit/secrets.toml`:

```toml
MISTRAL_API_KEY = "your-key"
```

Start the app:

```bash
streamlit run app.py
```

## Vessel master list

The app loads vessel names from `vessels.csv`. Keep the first-column header as `VESSEL`. If the file is missing, the current list is also included as an internal fallback.

Vessel names are used in the output filename, audit workbook, and prepared email. They are not written into the import template.

## Automatic sub-machinery detection

The extraction model reads the heading above each spare-parts table and returns:

- detected sub-machinery,
- table title,
- section start page,
- exact spare-part source page, and
- confidence.

The app groups repeated headings into editable proposals. Review the proposal name, code, first/last page, part count, confidence, and detected variants before applying assignments.

## Large manuals

For very large PDFs, process ranges such as `1-100`, `101-200`, and `201-300`. Enable **Append to current review table** after the first range. Download the audit workbook regularly because Streamlit session data is held in memory.

## Validation

The app blocks export when required machinery or spare-part fields are missing, machinery names do not match the machinery sheet, units are invalid, quantities are nonnumeric, or possible duplicates remain unresolved.

Always test a small import batch before production use.
