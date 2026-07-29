# Spare Parts OCR Import Builder

**Current build: 4.5**

A Streamlit application that converts scanned machinery manuals into structured machinery, sub-machinery, and spare-parts records, then writes approved data into the required Excel import template.

The application is designed for an automated, exception-only workflow: source drawing/table codes are used to connect each spare part to the correct sub-machinery, while users review only missing, conflicting, or low-confidence records.

## Main features

- Shared Gmail one-time-password access gate
- Multiple PDF jobs in one Streamlit session
- Separate vessel assignments and machinery details per PDF
- Direct Mistral OCR integration
- Automated sub-machinery detection from drawing/table sections
- Exact source-code-based matching between sub-machineries and spare parts
- English-only uppercase sub-machinery and spare-part names
- PDF-derived maker and model detection
- Automatic continuation-page merging
- Exception-only review workflow
- Content-aware UI column widths
- Import workbook, audit workbook, email draft, and multi-document ZIP export

## Workflow

1. Request a one-time access code sent to the shared Performance mailbox.
2. Upload one or more scanned PDF manuals.
3. Select the active PDF and assign its vessel or vessels.
4. Enter the required main-machinery details.
5. Run OCR and structured extraction.
6. Review only detected exceptions when necessary.
7. Generate the import workbook and audit package.

Each PDF keeps its own vessel assignment, machinery details, OCR output, review state, and export files.

## Automated sub-machinery matching

Each detected section is keyed internally by its exact drawing or table code.

Example:

```text
SOURCE CODE: 1-999-010
ENGLISH NAME: CRANK CASE
EXPORTED NAME: CRANK CASE (1-999-010)
```

The exact exported sub-machinery name is copied into the spare-parts `MACHINERY` column. This creates the connection between the machinery sheet and its corresponding spare-parts rows.

Continuation pages with the same source code remain under the same sub-machinery. A repeated table header does not create a duplicate sub-machinery.

## Sub-machinery field mapping

| Import field | Source or rule |
|---|---|
| `CODE` | Drawing/table code from the PDF |
| `NAME` | English uppercase section name plus source code in parentheses |
| `MAKER` | Maker detected for that section |
| `MODEL` | Model detected for that section |
| `TYPE` | Extracted when clearly available |
| `INSTR.BOOK` | Initial PDF page where the section begins |
| `SPECIFICATIONS` | Original PDF filename |
| `MCH_TP(M/S/U)` | `SubMachinery` |

Example:

```text
CODE: 1-999-010
NAME: CRANK CASE (1-999-010)
MAKER: MACGREGOR
MODEL: L 350
INSTR.BOOK: PDF PAGE.2
SPECIFICATIONS: PDF FILE: Technical Documentation L350.pdf
MCH_TP(M/S/U): SubMachinery
```

A maker or model detected specifically for a sub-machinery overrides the manually entered main-machinery values. The main-machinery values are used only as fallback.

## Spare-part field mapping

| Import field | Source or rule |
|---|---|
| `MACHINERY` | Exact exported sub-machinery `NAME` |
| `PART NO` | Same value as the extracted identification code |
| `DESCRIPTION` | English spare-part name in uppercase |
| `CODE` | `Ident-Nr.`, `Ident-No.`, or equivalent identification code |
| `ITEM NO` | Drawing position/reference number when available |
| `UNIT` | Defaults to `PCS` |
| `QNT` | Extracted numeric quantity |

Example:

```text
MACHINERY: CRANK CASE (1-999-010)
PART NO: 10.10.10.40
DESCRIPTION: SLIDE BEARING
CODE: 10.10.10.40
ITEM NO: 2
UNIT: PCS
QNT: 1
```

A separate manufacturer `Part No.` printed in the source is not populated automatically. Users may still enter or replace `PART NO` manually during review.

The first sequential drawing-position column is treated as `ITEM NO`, even when the source labels that column as `Part-No.`.

## Language normalization

- Existing English wording is preferred.
- Text is translated only when no English wording is available.
- Sub-machinery names are converted to English uppercase.
- Spare-part descriptions are converted to English uppercase.
- Codes, standards, dimensions, model numbers, and stage references are preserved.

## Review philosophy

The application automatically marks records ready when:

- a valid source section code is identified;
- the sub-machinery exists in the machinery list;
- the spare-part identification code or item number is available;
- the English description is confidently isolated;
- the quantity and unit are valid;
- no conflicting duplicate exists.

Only exceptions should remain in the review queue, including:

- missing section codes;
- conflicting section identities;
- missing or ambiguous English descriptions;
- conflicting maker/model values;
- unresolved duplicate rows;
- missing required identifiers;
- manually excluded records.

Manual changes are preserved when additional page ranges are appended.

## Large manuals

For large PDFs, process ranges such as:

```text
1-100
101-200
201-300
```

Enable **Append to this document's review table** after the first range.

The application preserves:

- the initial page of each section;
- previously approved sub-machinery assignments;
- manual spare-part corrections;
- manually entered `PART NO` values;
- previously extracted rows from the same document.

## User interface tables

The application estimates column widths using the heading and actual data length.

- Code, quantity, page, unit, and item-number columns remain compact.
- Machinery and description columns receive more space.
- Long warning and description columns are capped so one value does not make the full table unusable.
- Editable and read-only tables use the same readability logic.

## Access control

The full application remains hidden until a valid one-time code is entered.

The code:

- is sent to the configured shared mailbox;
- contains six digits;
- expires after the configured period;
- can be used only once;
- is tied to the browser session that requested it;
- is stored only as a salted cryptographic digest;
- is protected by resend cooldowns, attempt limits, and temporary lockout.

## Included files

- `app.py` — Streamlit interface, multi-document workflow, OTP gate, review, and exports
- `tools.py` — OCR, extraction, normalization, source-code matching, validation, and Excel generation
- `vessels.csv` — searchable vessel master list
- `Spare parts template last version.xlsx` — required import template
- `requirements.txt` — Python dependencies
- `README.md` — project instructions

## Local setup

Use Python 3.12.

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

macOS/Linux:

```bash
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Create:

```text
.streamlit/secrets.toml
```

Example:

```toml
MISTRAL_API_KEY = "your_mistral_api_key"

OTP_EMAIL = "perftechnomar@gmail.com"
GMAIL_SMTP_USER = "perftechnomar@gmail.com"
GMAIL_APP_PASSWORD = "your_16_character_google_app_password"

OTP_EXPIRY_MINUTES = 5
OTP_RESEND_SECONDS = 60
OTP_GLOBAL_RESEND_SECONDS = 30
OTP_MAX_ATTEMPTS = 5
OTP_LOCK_MINUTES = 10
AUTH_SESSION_HOURS = 8

OTP_SMTP_HOST = "smtp.gmail.com"
OTP_SMTP_PORT = 465
```

Do not commit real API keys or email credentials to GitHub.

Start the application:

```bash
streamlit run app.py
```

## Streamlit Cloud deployment

1. Upload or replace these files together:
   - `app.py`
   - `tools.py`
   - `requirements.txt`
2. Keep these files beside `app.py`:
   - `Spare parts template last version.xlsx`
   - `vessels.csv`
3. Add the Mistral and Gmail values under:
   - **Manage app → Settings → Secrets**
4. Save the secrets and reboot the application.
5. Confirm the login screen shows:

```text
Restricted access · Build 4.5
```

6. After login, confirm the main page shows:

```text
Build 4.5 — source-code matching, English uppercase normalization, and exception-only review
```

## Recommended validation after deployment

When Mistral usage is available:

1. Process pages `1-10` from a representative manual.
2. Confirm sub-machinery names follow:

```text
ENGLISH NAME (SOURCE CODE)
```

3. Confirm each spare-parts `MACHINERY` value exactly equals the corresponding machinery-sheet `NAME`.
4. Confirm the identification code is copied to both `CODE` and `PART NO`.
5. Confirm drawing positions are stored under `ITEM NO`.
6. Confirm `INSTR.BOOK` contains the section's initial PDF page.
7. Confirm `SPECIFICATIONS` contains the original PDF filename.
8. Confirm only genuine exceptions appear under **Needs correction**.
9. Test a small import batch before production use.

## Data handling

Uploaded pages are sent to the configured Mistral service when OCR or structured extraction runs.

Uploaded PDFs and active review data are stored temporarily for the Streamlit session. Download the audit workbook regularly when processing large manuals.

Use the application only for documents approved for external AI processing.
