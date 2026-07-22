from __future__ import annotations

import copy
import hashlib
import io
import tempfile
import uuid
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from tools import (
    PAGE_FILTER_MODES,
    REVIEW_COLUMNS,
    SUBMACHINERY_REVIEW_COLUMNS,
    UNIT_OPTIONS,
    add_manual_submachinery_candidate,
    apply_submachinery_assignments,
    build_audit_workbook,
    build_submachinery_candidates,
    build_workbook,
    classify_ocr_pages,
    empty_review_dataframe,
    empty_submachinery_review_dataframe,
    extract_spare_parts_from_markdown_tables,
    extract_spare_parts_with_ai,
    included_submachinery_rows,
    machinery_rows_from_main_and_additional,
    merge_review_dataframes,
    merge_submachinery_candidates,
    ocr_document_url,
    ocr_image_bytes,
    ocr_image_url,
    ocr_pdf_bytes,
    parse_page_spec,
    pdf_page_count,
    recalculate_review_status,
    rows_to_review_dataframe,
    safe_filename,
    validate_machinery_dataframe,
)


APP_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE_PATH = APP_DIR / "Spare parts template last version.xlsx"
APP_VERSION = "4.1"

DEFAULT_VESSEL_PATH = APP_DIR / "vessels.csv"

FALLBACK_VESSELS = [
    'AGIOS DIMITRIOS',
    'AFRICAN QUEEN',
    'ANTHEA Y',
    'ATETI',
    'ATHENA I',
    'BEATRICE',
    'BREMERHAVEN EXPRESS',
    'CAPTAIN THANASIS I',
    'CHRISTINAB',
    'CMA CGM ALCAZAR',
    'CMA CGM AMERICA',
    'CMA CGM JAMAICA',
    'CMA CGM SAMBHAR',
    'CMA CGM THALASSA',
    'COLOMBIA EXPRESS',
    'CONSTANTINOS P II',
    'COSTA RICA EXPRESS',
    'CYPRESS',
    'CZECH',
    'DARLEAKAY',
    'DOLPHIN II',
    'ELENI T',
    'EPAMINONDAS',
    'FRIEDERIKE',
    'GSL ALEXANDRA',
    'GSL ALICE',
    'GSL ARCADIA',
    "GSL CHATEAU D'IF",
    'GSL CHLOE',
    'GSL CHRISTEL ELISABETH',
    'GSL CHRISTEN',
    'GSL DOROTHEA',
    'GSL EFFIE',
    'GSL ELEFTHERIA',
    'GSL ELENI',
    'GSL ELIZABETH',
    'GSL GRANIA',
    'GSL KALLIOPI',
    'GSL KITHIRA',
    'GSL LALO',
    'GSL LYDIA',
    'GSL MAMITSA',
    'GSL MAREN',
    'GSL MARIA',
    'GSL MELINA',
    'GSL MELITA',
    'GSL MERCER',
    'GSL MYNY',
    'GSL NICOLETTA',
    'GSL NINGBO',
    'GSL ROSSI',
    'GSL SOFIA',
    'GSL SUSAN',
    'GSL SYROS',
    'GSL TEGEA',
    'GSL TINOS',
    'GSL TRIPOLI',
    'GSL VALERIE',
    'GSL VINIA',
    'GSL VIOLETTA',
    'IAN H',
    'ISTANBUL EXPRESS',
    'JAMAICA EXPRESS',
    'JULIE',
    'KACEY',
    'KOI',
    'KOSTAS K',
    'KUMASI',
    'LINDSAYLOU',
    'LOTUS A',
    'MAIRA',
    'MANET',
    'MARIA Y',
    'MARIANNA I',
    'MARINO',
    'MELINA',
    'MELINDA',
    'MEXICO EXPRESS',
    'MOON',
    'MSC QINGDAO',
    'MSC ROMA',
    'MSC TIANJIN',
    'MYNY',
    'NEWYORKER',
    'NICARAGUA EXPRESS',
    'NIKOLAS',
    'NIKOLAS XL',
    'ORCA I',
    'PANAMA EXPRESS',
    'SPYROS V',
    'STAMATIS B',
    'SYDNEY EXPRESS',
    'TINA I',
    'TONSBERG',
    'TORRANCE',
    'ZIM NORFOLK',
    'ZIM XIAMEN',
    'ZOI',
    'ZOI XL',
]


def load_vessel_options() -> list[str]:
    if DEFAULT_VESSEL_PATH.exists():
        try:
            vessel_frame = pd.read_csv(DEFAULT_VESSEL_PATH, dtype=str)
            if not vessel_frame.empty:
                first_column = vessel_frame.columns[0]
                values = [
                    str(value).strip()
                    for value in vessel_frame[first_column].dropna().tolist()
                    if str(value).strip()
                ]
                if values:
                    return sorted(dict.fromkeys(values), key=str.upper)
        except Exception:
            pass
    return sorted(dict.fromkeys(FALLBACK_VESSELS), key=str.upper)


VESSEL_OPTIONS = load_vessel_options()

DEFAULT_PAGE_FILTER = next(
    (mode for mode in PAGE_FILTER_MODES if "conservative" in mode.lower()),
    PAGE_FILTER_MODES[0],
)

PROCESSING_PRESETS = {
    "Balanced": {
        "description": (
            "Recommended for most manuals. Good balance between speed, stability, "
            "and extraction accuracy."
        ),
        "structure_mode": "AI JSON extraction (recommended)",
        "page_filter_mode": DEFAULT_PAGE_FILTER,
        "extraction_model": "mistral-small-latest",
        "ocr_pages_per_request": 25,
        "extraction_pages_per_batch": 3,
        "extraction_max_chars": 12000,
        "default_unit": "PCS",
    },
    "Fast": {
        "description": (
            "For clean, consistent scans. Processes larger batches, so it is faster "
            "but may need more review."
        ),
        "structure_mode": "AI JSON extraction (recommended)",
        "page_filter_mode": DEFAULT_PAGE_FILTER,
        "extraction_model": "mistral-small-latest",
        "ocr_pages_per_request": 35,
        "extraction_pages_per_batch": 4,
        "extraction_max_chars": 16000,
        "default_unit": "PCS",
    },
    "Careful": {
        "description": (
            "For poor scans, complex layouts, or repeated recovery messages. Uses "
            "small batches for maximum stability."
        ),
        "structure_mode": "AI JSON extraction (recommended)",
        "page_filter_mode": DEFAULT_PAGE_FILTER,
        "extraction_model": "mistral-small-latest",
        "ocr_pages_per_request": 12,
        "extraction_pages_per_batch": 1,
        "extraction_max_chars": 8000,
        "default_unit": "PCS",
    },
}

st.set_page_config(
    page_title="Spare Parts OCR Import Builder",
    page_icon="📄",
    layout="wide",
)


# ---------------------------------------------------------------------------
# Session state
# ---------------------------------------------------------------------------


def initialize_state() -> None:
    defaults = {
        "extracted_pages": [],
        "page_classification": pd.DataFrame(),
        "extraction_log": [],
        "spare_review": empty_review_dataframe(),
        "submachinery_review": empty_submachinery_review_dataframe(),
        "selected_vessels": [],
        "additional_vessels_text": "",
        "output": None,
        "output_name": "spare_parts.xlsx",
        "editor_version": 0,
        "main_code": "",
        "main_name": "",
        "main_maker": "",
        "main_model": "",
        "main_type": "",
        "main_instruction_book": "",
        "main_specifications": "",
        "auto_instruction_book_source": "",
        "processing_preset": "Balanced",
        "setting_structure_mode": PROCESSING_PRESETS["Balanced"]["structure_mode"],
        "setting_page_filter_mode": PROCESSING_PRESETS["Balanced"]["page_filter_mode"],
        "setting_extraction_model": PROCESSING_PRESETS["Balanced"]["extraction_model"],
        "setting_ocr_pages_per_request": PROCESSING_PRESETS["Balanced"]["ocr_pages_per_request"],
        "setting_extraction_pages_per_batch": PROCESSING_PRESETS["Balanced"]["extraction_pages_per_batch"],
        "setting_extraction_max_chars": PROCESSING_PRESETS["Balanced"]["extraction_max_chars"],
        "setting_default_unit": PROCESSING_PRESETS["Balanced"]["default_unit"],
        "setting_extra_prompt": "",
        "submachinery_editor_version": 0,
        "review_filter": "Needs correction",
        "review_sort": "Issues first",
        "review_confidence_threshold": 0.75,
        "source_page_lookup": 1,
        "prepared_email_subject": "",
        "prepared_email_body": "",
        "active_page_spec": "all",
        "append_results": False,
        "document_jobs": {},
        "session_token": uuid.uuid4().hex,
        "loaded_job_id": "",
        "active_document_id": "",
        "active_document_selector": "",
        "multi_package_output": None,
        "multi_package_name": "multi_document_import_package.zip",
        "multi_package_report": [],
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


JOB_STATE_FIELDS = [
    "extracted_pages",
    "page_classification",
    "extraction_log",
    "spare_review",
    "submachinery_review",
    "selected_vessels",
    "additional_vessels_text",
    "output",
    "output_name",
    "editor_version",
    "main_code",
    "main_name",
    "main_maker",
    "main_model",
    "main_type",
    "main_instruction_book",
    "main_specifications",
    "auto_instruction_book_source",
    "submachinery_editor_version",
    "review_filter",
    "review_sort",
    "review_confidence_threshold",
    "source_page_lookup",
    "prepared_email_subject",
    "prepared_email_body",
    "active_page_spec",
    "append_results",
]

JOB_STORAGE_DIR = Path(tempfile.gettempdir()) / "spare_parts_builder_jobs"
JOB_STORAGE_DIR.mkdir(parents=True, exist_ok=True)


def _clone_state_value(value):
    if isinstance(value, pd.DataFrame):
        return value.copy(deep=True)
    return copy.deepcopy(value)


def _empty_job_state(file_name: str, pdf_path: str, file_hash: str, size_bytes: int) -> dict:
    return {
        "job_id": file_hash[:16],
        "file_hash": file_hash,
        "file_name": file_name,
        "pdf_path": pdf_path,
        "size_bytes": int(size_bytes),
        "extracted_pages": [],
        "page_classification": pd.DataFrame(),
        "extraction_log": [],
        "spare_review": empty_review_dataframe(),
        "submachinery_review": empty_submachinery_review_dataframe(),
        "selected_vessels": [],
        "additional_vessels_text": "",
        "output": None,
        "output_name": "spare_parts.xlsx",
        "editor_version": 0,
        "main_code": "",
        "main_name": "",
        "main_maker": "",
        "main_model": "",
        "main_type": "",
        "main_instruction_book": file_name,
        "main_specifications": "",
        "auto_instruction_book_source": file_name,
        "submachinery_editor_version": 0,
        "review_filter": "Needs correction",
        "review_sort": "Issues first",
        "review_confidence_threshold": 0.75,
        "source_page_lookup": 1,
        "prepared_email_subject": "",
        "prepared_email_body": "",
        "active_page_spec": "all",
        "append_results": False,
    }


def save_loaded_job_state() -> None:
    job_id = str(st.session_state.get("loaded_job_id", ""))
    jobs = st.session_state.get("document_jobs", {})
    if not job_id or job_id not in jobs:
        return
    job = jobs[job_id]
    for field in JOB_STATE_FIELDS:
        if field in st.session_state:
            job[field] = _clone_state_value(st.session_state[field])
    jobs[job_id] = job
    st.session_state.document_jobs = jobs


def load_document_job(job_id: str) -> None:
    jobs = st.session_state.get("document_jobs", {})
    if job_id not in jobs:
        return
    job = jobs[job_id]
    for field in JOB_STATE_FIELDS:
        if field in job:
            st.session_state[field] = _clone_state_value(job[field])
    st.session_state.loaded_job_id = job_id
    st.session_state.active_document_id = job_id


def register_uploaded_pdfs(uploaded_files) -> None:
    if not uploaded_files:
        return
    jobs = st.session_state.get("document_jobs", {})
    for uploaded in uploaded_files:
        data = uploaded.getvalue()
        file_hash = hashlib.sha256(data).hexdigest()
        job_id = file_hash[:16]
        session_token = str(st.session_state.get("session_token", "session"))
        pdf_path = JOB_STORAGE_DIR / f"{session_token}_{file_hash}.pdf"
        if not pdf_path.exists():
            pdf_path.write_bytes(data)
        if job_id not in jobs:
            jobs[job_id] = _empty_job_state(
                file_name=uploaded.name,
                pdf_path=str(pdf_path),
                file_hash=file_hash,
                size_bytes=len(data),
            )
        else:
            jobs[job_id]["file_name"] = uploaded.name
            jobs[job_id]["pdf_path"] = str(pdf_path)
            jobs[job_id]["size_bytes"] = len(data)
    st.session_state.document_jobs = jobs


def remove_document_job(job_id: str) -> None:
    jobs = st.session_state.get("document_jobs", {})
    job = jobs.pop(job_id, None)
    if job:
        pdf_path = Path(str(job.get("pdf_path", "")))
        try:
            if pdf_path.exists():
                pdf_path.unlink()
        except OSError:
            pass
    st.session_state.document_jobs = jobs
    remaining = list(jobs)
    next_job = remaining[0] if remaining else ""
    st.session_state.loaded_job_id = ""
    st.session_state.active_document_id = next_job
    st.session_state.active_document_selector = next_job
    if next_job:
        load_document_job(next_job)


class StoredPdf:
    def __init__(self, file_name: str, pdf_path: str):
        self.name = file_name
        self._path = Path(pdf_path)

    def getvalue(self) -> bytes:
        return self._path.read_bytes()


def active_document_job() -> dict | None:
    job_id = str(st.session_state.get("loaded_job_id", ""))
    return st.session_state.get("document_jobs", {}).get(job_id)


def _job_vessels(job: dict) -> list[str]:
    selected = [str(value).strip() for value in job.get("selected_vessels", []) if str(value).strip()]
    additional_text = str(job.get("additional_vessels_text", ""))
    additional = [
        item.strip()
        for item in additional_text.replace(";", "\n").replace(",", "\n").splitlines()
        if item.strip()
    ]
    return list(dict.fromkeys(selected + additional))


def _job_main_row(job: dict) -> dict[str, str]:
    return {
        "CODE": str(job.get("main_code", "")),
        "NAME": str(job.get("main_name", "")),
        "MAKER": str(job.get("main_maker", "")),
        "MODEL": str(job.get("main_model", "")),
        "TYPE": str(job.get("main_type", "")),
        "INSTR.BOOK": str(job.get("main_instruction_book", "")),
        "SPECIFICATIONS": str(job.get("main_specifications", "")),
        "MCH_TP(M/S/U)": "Main Machinery",
    }


def _job_machinery_frame(job: dict) -> pd.DataFrame:
    sub_frame = included_submachinery_rows(job.get("submachinery_review", empty_submachinery_review_dataframe()))
    return machinery_rows_from_main_and_additional(_job_main_row(job), sub_frame)


def _job_status(job: dict) -> tuple[str, int, int]:
    review = job.get("spare_review", empty_review_dataframe())
    rows = len(review) if isinstance(review, pd.DataFrame) else 0
    sub_frame = job.get("submachinery_review", empty_submachinery_review_dataframe())
    sub_count = len(sub_frame) if isinstance(sub_frame, pd.DataFrame) else 0
    if job.get("output"):
        return "Export created", rows, sub_count
    if rows:
        return "Review in progress", rows, sub_count
    if job.get("extracted_pages"):
        return "OCR complete", rows, sub_count
    return "Not processed", rows, sub_count


def get_secret(name: str) -> str:
    try:
        return str(st.secrets[name]).strip()
    except Exception:
        return ""


def apply_processing_preset() -> None:
    preset_name = st.session_state.get("processing_preset", "Balanced")
    preset = PROCESSING_PRESETS.get(preset_name, PROCESSING_PRESETS["Balanced"])
    st.session_state.setting_structure_mode = preset["structure_mode"]
    st.session_state.setting_page_filter_mode = preset["page_filter_mode"]
    st.session_state.setting_extraction_model = preset["extraction_model"]
    st.session_state.setting_ocr_pages_per_request = preset["ocr_pages_per_request"]
    st.session_state.setting_extraction_pages_per_batch = preset["extraction_pages_per_batch"]
    st.session_state.setting_extraction_max_chars = preset["extraction_max_chars"]
    st.session_state.setting_default_unit = preset["default_unit"]


initialize_state()
save_loaded_job_state()

st.title("📄 Spare Parts OCR Import Builder")


# ---------------------------------------------------------------------------
# Sidebar settings
# ---------------------------------------------------------------------------

with st.sidebar:
    with st.expander("📖 Instructions & Help", expanded=False):
        st.markdown(
            """
### Quick start

1. Under **Documents**, upload one or more PDF manuals and select the active document.
2. Open **1. Machinery**, assign vessel(s) to that PDF, and complete its main machinery fields.
3. Keep **Balanced** processing mode for normal use, then run **2. OCR**.
4. Open **3. Sub-machineries** to review the automatically detected table headings. Each proposal shows its source-page range and number of linked spare parts.
5. Apply the approved sub-machinery assignments, then open **4. Review spare parts**.
6. Use the filters, source-page columns, and quick page lookup to correct only the rows that need attention.
7. Open **5. Export** to create the active document workbook or a ZIP package for every ready document.

### Multiple documents and vessel assignment

- Every uploaded PDF keeps its own machinery, vessel, OCR, review, and export state.
- Switch the **Active document** in the sidebar to configure another PDF.
- Select one or more vessels from the searchable list for the active PDF.
- Vessel names are used in the export filename, audit workbook, and email draft.
- Vessel names are **not written into the import template**, because vessel assignment is completed during the ERP process.
- Use **Additional vessel(s)** only when a vessel is not yet available in the master list.

### Automatic sub-machinery detection

The AI reads the title above each spare-parts table and proposes sub-machineries automatically in the dedicated **3. Sub-machineries** tab.

- **FIRST PAGE / LAST PAGE:** where the detected table section appears.
- **PARTS FOUND:** number of spare-part rows linked to the proposal.
- **CONFIDENCE:** average extraction confidence for that detected section.
- **VARIANTS:** different spellings found in the manual.
- Uncheck **INCLUDE** to reject a false detection, or edit **NAME** and **CODE** before applying assignments.
- Edit several cells without interruption, then select **Save sub-machinery changes**.
- Use **Apply approved assignments to spare parts** after saving, renaming, excluding, or merging proposals.

### Processing modes

**Balanced — recommended**  
Best for most manuals. Good balance of speed, stability, and extraction accuracy.

**Fast**  
For clean, consistent scans and regular tables. Larger batches are faster but may need more review.

**Careful**  
For poor scans, complex layouts, OCR timeouts, or repeated recovery messages. Smaller batches are slower but more stable.

### Advanced Mistral settings

Normal users only need a processing mode. Open **Advanced Mistral settings** for difficult manuals. Re-selecting a mode restores that mode's default values.

### Review dashboard

- **Needs correction** opens by default.
- Filter by low confidence, unassigned sub-machinery, ready, excluded, or all rows.
- Sort by issues, confidence, source page, section start page, sub-machinery, part number, or description.
- **SUB-MACHINERY** is the approved machinery record assigned to that spare-part row. It may show the main machinery only when no sub-machinery applies.
- **SOURCE PAGE** is the exact page containing the spare-part row.
- **SECTION START PAGE** is where the detected sub-machinery/table section began.
- Use **Source page quick lookup** to display the OCR text for a page while the original PDF is open on a second monitor.
- Corrections are saved to the full dataset even when only a filtered subset is visible.

### Large manuals

For very large books, process page ranges such as `1-100`, `101-200`, and so on. Each PDF is processed independently. Enable **Append to current review table** after the first range. Download the audit workbook regularly because an app restart clears in-memory data.

### Data handling

Uploaded pages are sent to the configured Mistral service when OCR or AI extraction runs. Use the tool only for documents approved for that processing.
            """
        )

    st.header("Documents")
    input_type = st.radio(
        "Choose input type",
        ["PDF", "Document URL", "Image", "Image URL"],
        index=0,
        help=(
            "Multi-document job management is available for PDF manuals. "
            "URL and image sources continue to use the current single workspace."
        ),
    )

    source_file = None
    document_url = ""
    image_url = ""
    page_spec = "all"
    append_results = False

    if input_type == "PDF":
        uploaded_pdf_files = st.file_uploader(
            "Upload one or more scanned PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            key="multi_pdf_uploader",
            help=(
                "Each unique PDF becomes an independent job. A PDF is OCR-processed once, "
                "even when it is assigned to several vessels."
            ),
        )
        register_uploaded_pdfs(uploaded_pdf_files)
        jobs = st.session_state.document_jobs

        if jobs:
            job_ids = list(jobs)
            if st.session_state.get("active_document_selector") not in jobs:
                st.session_state.active_document_selector = job_ids[0]

            selected_job_id = st.selectbox(
                "Active document",
                options=job_ids,
                key="active_document_selector",
                format_func=lambda value: jobs[value]["file_name"],
                help="Switch documents without losing the OCR, review, vessel, or export state of the other jobs.",
            )

            if st.session_state.get("loaded_job_id") != selected_job_id:
                save_loaded_job_state()
                load_document_job(selected_job_id)
                st.rerun()

            active_job = jobs[selected_job_id]
            source_file = StoredPdf(active_job["file_name"], active_job["pdf_path"])
            page_spec = st.text_input(
                "Pages to process for active document",
                key="active_page_spec",
                help="Examples: all, 1-20, 25, 30-35.",
            )
            append_results = st.checkbox(
                "Append to this document's review table",
                key="append_results",
                help="Use when processing additional page ranges from the same PDF.",
            )
            active_status, active_rows, active_subs = _job_status(active_job)
            st.caption(
                f"{active_status} · {active_rows} spare-part row(s) · "
                f"{active_subs} sub-machinery proposal(s)"
            )
            st.button(
                "Remove active document",
                use_container_width=True,
                on_click=remove_document_job,
                args=(selected_job_id,),
            )
        else:
            st.info("Upload one or more PDFs to create document jobs.")
    elif input_type == "Document URL":
        document_url = st.text_input(
            "Document URL",
            help="Enter a direct, publicly accessible PDF URL.",
        )
        append_results = st.checkbox(
            "Append to current review table",
            value=False,
            help="Single-workspace behavior for URL sources.",
        )
    elif input_type == "Image":
        source_file = st.file_uploader(
            "Upload an image",
            type=["png", "jpg", "jpeg"],
            help="Use this for a single scanned page or photograph.",
        )
        append_results = st.checkbox(
            "Append to current review table",
            value=False,
            help="Single-workspace behavior for image sources.",
        )
    else:
        image_url = st.text_input(
            "Image URL",
            help="Enter a direct, publicly accessible image URL.",
        )
        append_results = st.checkbox(
            "Append to current review table",
            value=False,
            help="Single-workspace behavior for image URL sources.",
        )

    st.divider()
    st.header("Processing mode")
    secret_api_key = get_secret("MISTRAL_API_KEY")
    if secret_api_key:
        entered_api_key = ""
    else:
        entered_api_key = st.text_input(
            "Mistral API key",
            type="password",
            help="For local testing only. Prefer .streamlit/secrets.toml for deployment.",
        )
    api_key = secret_api_key or entered_api_key

    st.caption(
        "Choose one mode. Balanced is recommended for normal use; Advanced settings "
        "remain available below for fine-tuning."
    )
    mode_columns = st.columns(3)
    for column, preset_name in zip(mode_columns, PROCESSING_PRESETS):
        selected = st.session_state.processing_preset == preset_name
        label = f"✓ {preset_name}" if selected else preset_name
        if column.button(
            label,
            key=f"preset_button_{preset_name.lower()}",
            use_container_width=True,
            help=PROCESSING_PRESETS[preset_name]["description"],
        ):
            st.session_state.processing_preset = preset_name
            apply_processing_preset()
            st.rerun()

    active_preset = PROCESSING_PRESETS.get(
        st.session_state.processing_preset, PROCESSING_PRESETS["Balanced"]
    )
    st.caption(f"**Active mode:** {st.session_state.processing_preset}")
    st.caption(active_preset["description"])

    with st.expander("Advanced Mistral settings", expanded=False):
        st.caption(
            "These are the active settings for the current run. Re-selecting a mode "
            "restores that mode's defaults."
        )

        structure_mode = st.selectbox(
            "Convert OCR text into rows",
            ["AI JSON extraction (recommended)", "Local markdown-table parser"],
            key="setting_structure_mode",
            help=(
                "AI JSON extraction handles irregular tables and wrapped descriptions. "
                "The local parser is faster but works best when OCR already produced clean Markdown tables."
            ),
        )
        page_filter_mode = st.selectbox(
            "Page filtering before AI extraction",
            PAGE_FILTER_MODES,
            key="setting_page_filter_mode",
            help=(
                "Conservative skips only obvious contents/revision/prose pages. Strict "
                "processes only strong parts-table candidates. Off processes every OCR page."
            ),
        )
        extraction_model = st.text_input(
            "Structured-extraction model",
            key="setting_extraction_model",
            disabled=structure_mode != "AI JSON extraction (recommended)",
            help="Recommended default: mistral-small-latest.",
        )
        ocr_pages_per_request = st.number_input(
            "PDF pages per OCR request",
            min_value=1,
            max_value=100,
            step=1,
            key="setting_ocr_pages_per_request",
            help="Lower values improve stability for poor scans or OCR timeouts.",
        )
        extraction_pages_per_batch = st.number_input(
            "OCR pages per structuring batch",
            min_value=1,
            max_value=20,
            step=1,
            key="setting_extraction_pages_per_batch",
            disabled=structure_mode != "AI JSON extraction (recommended)",
            help="Lower values reduce malformed or truncated JSON responses.",
        )
        extraction_max_chars = st.number_input(
            "Maximum OCR characters per AI batch",
            min_value=2000,
            max_value=30000,
            step=1000,
            key="setting_extraction_max_chars",
            disabled=structure_mode != "AI JSON extraction (recommended)",
            help=(
                "Smaller batches reduce malformed/truncated JSON. Failed batches are "
                "also divided automatically into smaller requests."
            ),
        )
        default_unit = st.selectbox(
            "Default spare-part unit",
            ["PCS", "SET", ""],
            key="setting_default_unit",
            help="PCS is the normal default. Use SET for kits or leave blank for manual review.",
        )
        extra_prompt = st.text_area(
            "Optional manual-specific extraction instructions",
            placeholder=(
                "Example: The first column is ITEM NO and the second column is PART NO. "
                "Ignore drawing dimensions and prices."
            ),
            height=100,
            key="setting_extra_prompt",
            disabled=structure_mode != "AI JSON extraction (recommended)",
            help="Add only rules that are specific to the current manual.",
        )

        st.markdown(
            f"""
**Current active values**

- Mode: `{st.session_state.processing_preset}`
- Page filtering: `{page_filter_mode}`
- Extraction method: `{structure_mode}`
- Model: `{extraction_model}`
- OCR pages/request: `{int(ocr_pages_per_request)}`
- Structuring pages/batch: `{int(extraction_pages_per_batch)}`
- Maximum characters/batch: `{int(extraction_max_chars)}`
- Default unit: `{default_unit or 'Blank'}`
            """
        )

    st.divider()
    st.header("Template")
    custom_template = st.file_uploader(
        "Optional replacement template",
        type=["xlsx"],
        help="Leave empty to use the template bundled with this app.",
        key="template_uploader",
    )
    if custom_template is not None:
        template_bytes = custom_template.getvalue()
        template_name = custom_template.name
        st.success(f"Using uploaded template: {template_name}")
    elif DEFAULT_TEMPLATE_PATH.exists():
        template_bytes = DEFAULT_TEMPLATE_PATH.read_bytes()
        template_name = DEFAULT_TEMPLATE_PATH.name
        st.success(f"Using bundled template: {template_name}")
    else:
        template_bytes = None
        template_name = ""
        st.error("Place the template beside app.py or upload it here.")

    if st.button(
        "Reset active document OCR and review data",
        use_container_width=True,
        help="Clears extracted pages, classifications, candidate rows and the generated workbook from this session.",
    ):
        st.session_state.extracted_pages = []
        st.session_state.page_classification = pd.DataFrame()
        st.session_state.extraction_log = []
        st.session_state.spare_review = empty_review_dataframe()
        st.session_state.submachinery_review = empty_submachinery_review_dataframe()
        st.session_state.output = None
        st.session_state.multi_package_output = None
        st.session_state.prepared_email_subject = ""
        st.session_state.prepared_email_body = ""
        st.session_state.editor_version += 1
        st.session_state.submachinery_editor_version += 1
        st.rerun()

    st.divider()
    with st.expander("ℹ️ About", expanded=False):
        st.markdown(
            f"""
**Spare Parts OCR Import Builder — v{APP_VERSION}**

**Workflow**  
Multiple PDFs → Per-document vessels and machinery → OCR → Review → Individual or package export

**Supported sources**  
Scanned PDFs, document URLs, images and image URLs.

**Important**  
The generated workbook should be tested with a small import batch before production use.
            """
        )


# ---------------------------------------------------------------------------
# Multi-document dashboard
# ---------------------------------------------------------------------------

if st.session_state.document_jobs:
    save_loaded_job_state()
    with st.expander("Document job dashboard", expanded=True):
        dashboard_rows = []
        for job_id, job in st.session_state.document_jobs.items():
            status, row_count, sub_count = _job_status(job)
            dashboard_rows.append(
                {
                    "ACTIVE": job_id == st.session_state.get("loaded_job_id"),
                    "DOCUMENT": job.get("file_name", ""),
                    "VESSELS": ", ".join(_job_vessels(job)) or "Not assigned",
                    "MAIN MACHINERY": job.get("main_name", "") or "Not entered",
                    "OCR PAGES": len(job.get("extracted_pages", [])),
                    "SUB-MACHINERIES": sub_count,
                    "SPARE-PART ROWS": row_count,
                    "STATUS": status,
                }
            )
        st.dataframe(
            pd.DataFrame(dashboard_rows),
            use_container_width=True,
            hide_index=True,
        )
        st.caption(
            "The same PDF may be assigned to many vessels without repeating OCR. "
            "Each PDF can also have a different vessel list."
        )


# ---------------------------------------------------------------------------
# Vessel, main machinery, and detected sub-machineries
# ---------------------------------------------------------------------------

# Use the uploaded PDF name as the initial instruction-book value, without
# overwriting a value the user has already entered.
if (
    input_type == "PDF"
    and source_file is not None
    and not st.session_state.main_instruction_book
    and st.session_state.auto_instruction_book_source != source_file.name
):
    st.session_state.main_instruction_book = source_file.name
    st.session_state.auto_instruction_book_source = source_file.name


def selected_vessel_names() -> list[str]:
    selected = [
        str(value).strip()
        for value in st.session_state.get("selected_vessels", [])
        if str(value).strip()
    ]
    additional_text = str(st.session_state.get("additional_vessels_text", ""))
    additional = [
        item.strip()
        for item in additional_text.replace(";", "\n").replace(",", "\n").splitlines()
        if item.strip()
    ]
    return list(dict.fromkeys(selected + additional))


def current_main_row() -> dict[str, str]:
    return {
        "CODE": st.session_state.main_code,
        "NAME": st.session_state.main_name,
        "MAKER": st.session_state.main_maker,
        "MODEL": st.session_state.main_model,
        "TYPE": st.session_state.main_type,
        "INSTR.BOOK": st.session_state.main_instruction_book,
        "SPECIFICATIONS": st.session_state.main_specifications,
        "MCH_TP(M/S/U)": "Main Machinery",
    }


def current_submachinery_rows() -> pd.DataFrame:
    return included_submachinery_rows(st.session_state.submachinery_review)


def current_machinery_frame() -> pd.DataFrame:
    return machinery_rows_from_main_and_additional(
        current_main_row(),
        current_submachinery_rows(),
    )


def main_machinery_is_ready() -> bool:
    required_keys = ("main_code", "main_name", "main_maker", "main_model")
    return bool(selected_vessel_names()) and all(
        str(st.session_state.get(key, "")).strip() for key in required_keys
    )


machinery_tab, input_tab, submachinery_tab, review_tab, export_tab = st.tabs(
    [
        "1. Machinery",
        "2. OCR",
        "3. Sub-machineries",
        "4. Review spare parts",
        "5. Export",
    ]
)

with machinery_tab:
    active_job = active_document_job()
    if active_job:
        st.caption(f"Active document: **{active_job['file_name']}**")
    st.subheader("Step 1 — Vessel assignment")
    st.multiselect(
        "Vessel(s) *",
        options=VESSEL_OPTIONS,
        key="selected_vessels",
        placeholder="Search and select one or more vessels",
        help=(
            "The vessel selection is used in the audit file, output filename, and "
            "email draft. It is not written into the import workbook."
        ),
    )
    with st.expander("Additional vessel(s) not in the list", expanded=False):
        st.text_area(
            "Enter one vessel per line",
            key="additional_vessels_text",
            height=90,
            help="Use only for vessels that are not yet present in vessels.csv.",
        )
    vessels = selected_vessel_names()
    if vessels:
        st.success(f"Selected vessel(s): {', '.join(vessels)}")
    else:
        st.warning("Select at least one vessel before running OCR.")

    st.subheader("Main machinery")
    st.info(
        "Enter the main machinery once. The app will detect the sub-machinery/table "
        "headings during OCR and propose the remaining machinery rows automatically."
    )

    row1 = st.columns(4)
    with row1[0]:
        st.text_input(
            "CODE *",
            key="main_code",
            placeholder="M/E",
            help="Required main-machinery code written to sheet 1.",
        )
    with row1[1]:
        st.text_input(
            "NAME *",
            key="main_name",
            placeholder="Main Engine",
            help="Required main-machinery name.",
        )
    with row1[2]:
        st.text_input(
            "MAKER *",
            key="main_maker",
            placeholder="MAN",
            help="Required manufacturer, normally found on the first pages/nameplate.",
        )
    with row1[3]:
        st.text_input(
            "MODEL *",
            key="main_model",
            placeholder="6S50MC-C",
            help="Required model, normally found on the first pages/nameplate.",
        )

    row2 = st.columns(3)
    with row2[0]:
        st.text_input(
            "TYPE",
            key="main_type",
            help="Optional type or variant from the manual/nameplate.",
        )
    with row2[1]:
        st.text_input(
            "INSTR.BOOK",
            key="main_instruction_book",
            help="The uploaded PDF filename is filled automatically when this is empty.",
        )
    with row2[2]:
        st.text_input(
            "SPECIFICATIONS",
            key="main_specifications",
            help="Optional technical specifications or distinguishing information.",
        )

    st.text_input(
        "MCH_TP(M/S/U)",
        value="Main Machinery",
        disabled=True,
        key="fixed_main_machinery_type",
    )



with submachinery_tab:
    active_job = active_document_job()
    if active_job:
        st.caption(f"Active document: **{active_job['file_name']}**")
    st.subheader("Step 3 — Detected sub-machineries")
    st.caption(
        "After OCR, the app groups the titles found above spare-parts tables. Review "
        "the proposed names, codes, source-page ranges, and variants before export."
    )

    # Normalize state after upgrades or older sessions.
    sub_frame = st.session_state.submachinery_review.copy()
    for column in SUBMACHINERY_REVIEW_COLUMNS:
        if column not in sub_frame.columns:
            sub_frame[column] = False if column == "INCLUDE" else ""
    if not sub_frame.empty:
        sub_frame = sub_frame[SUBMACHINERY_REVIEW_COLUMNS]
        sub_frame["MCH_TP(M/S/U)"] = "SubMachinery"
    st.session_state.submachinery_review = sub_frame

    sub_action_cols = st.columns([1.25, 1.5, 1.5, 1.1])
    with sub_action_cols[0]:
        if st.button("Add manual sub-machinery", use_container_width=True):
            st.session_state.submachinery_review = add_manual_submachinery_candidate(
                st.session_state.submachinery_review,
                current_main_row(),
            )
            st.session_state.submachinery_editor_version += 1
            st.rerun()
    with sub_action_cols[1]:
        if st.button(
            "Fill missing details from main",
            use_container_width=True,
            disabled=st.session_state.submachinery_review.empty,
        ):
            frame = st.session_state.submachinery_review.copy()
            defaults = current_main_row()
            for column in ("MAKER", "MODEL", "INSTR.BOOK"):
                blank = frame[column].astype(str).str.strip().eq("")
                frame.loc[blank, column] = defaults[column]
            frame["MCH_TP(M/S/U)"] = "SubMachinery"
            st.session_state.submachinery_review = frame
            st.session_state.submachinery_editor_version += 1
            st.rerun()
    with sub_action_cols[2]:
        if st.button(
            "Apply approved assignments to spare parts",
            use_container_width=True,
            disabled=(
                st.session_state.submachinery_review.empty
                or st.session_state.spare_review.empty
            ),
            help=(
                "Applies included proposal names to linked spare-part rows. Run this "
                "after renaming or excluding sub-machineries."
            ),
        ):
            st.session_state.spare_review = apply_submachinery_assignments(
                st.session_state.spare_review,
                st.session_state.submachinery_review,
                st.session_state.main_name,
                overwrite_auto_assignments=True,
            )
            st.session_state.editor_version += 1
            st.success("Approved sub-machinery assignments were applied.")
    with sub_action_cols[3]:
        if st.button(
            "Clear proposals",
            use_container_width=True,
            disabled=st.session_state.submachinery_review.empty,
        ):
            st.session_state.submachinery_review = empty_submachinery_review_dataframe()
            st.session_state.submachinery_editor_version += 1
            st.rerun()

    if st.session_state.submachinery_review.empty:
        st.info(
            "No sub-machineries detected yet. Complete the main machinery, upload the "
            "manual, and run OCR. Proposals will appear here automatically."
        )
    else:
        candidate_frame = st.session_state.submachinery_review
        candidate_metrics = st.columns(4)
        candidate_metrics[0].metric("Proposals", len(candidate_frame))
        candidate_metrics[1].metric(
            "Included", int(candidate_frame["INCLUDE"].astype(bool).sum())
        )
        candidate_metrics[2].metric(
            "Linked parts",
            int(pd.to_numeric(candidate_frame["PARTS FOUND"], errors="coerce").fillna(0).sum()),
        )
        candidate_metrics[3].metric(
            "Low confidence",
            int(
                (
                    pd.to_numeric(candidate_frame["CONFIDENCE"], errors="coerce").fillna(0)
                    < 0.75
                ).sum()
            ),
        )

        active_job_id = str(st.session_state.get("loaded_job_id", "single"))
        form_key = (
            f"submachinery_edit_form_{active_job_id}_"
            f"{st.session_state.submachinery_editor_version}"
        )
        editor_key = (
            f"submachinery_editor_{active_job_id}_"
            f"{st.session_state.submachinery_editor_version}"
        )

        st.info(
            "Edit as many cells as needed, including in full-screen view. Cell edits "
            "are held without rerunning the app. Exit full-screen and select "
            "**Save sub-machinery changes** once when finished."
        )

        with st.form(form_key, clear_on_submit=False, border=False):
            edited_submachineries = st.data_editor(
                candidate_frame,
                key=editor_key,
                num_rows="fixed",
                use_container_width=True,
                hide_index=True,
                height=520,
                disabled=[
                    "MCH_TP(M/S/U)",
                    "FIRST PAGE",
                    "LAST PAGE",
                    "PARTS FOUND",
                    "CONFIDENCE",
                    "VARIANTS",
                    "ORIGIN",
                ],
                column_config={
                    "INCLUDE": st.column_config.CheckboxColumn(
                        "INCLUDE",
                        help="Only included rows are written to the machinery sheet.",
                        default=True,
                    ),
                    "CODE": st.column_config.TextColumn("CODE", width="small"),
                    "NAME": st.column_config.TextColumn(
                        "NAME",
                        width="large",
                        help="Canonical sub-machinery name used by linked spare parts.",
                    ),
                    "MAKER": st.column_config.TextColumn("MAKER", width="medium"),
                    "MODEL": st.column_config.TextColumn("MODEL", width="medium"),
                    "TYPE": st.column_config.TextColumn("TYPE", width="medium"),
                    "INSTR.BOOK": st.column_config.TextColumn("INSTR.BOOK", width="medium"),
                    "SPECIFICATIONS": st.column_config.TextColumn(
                        "SPECIFICATIONS", width="medium"
                    ),
                    "MCH_TP(M/S/U)": st.column_config.TextColumn(
                        "MCH_TP(M/S/U)", width="small"
                    ),
                    "FIRST PAGE": st.column_config.NumberColumn(
                        "FIRST PAGE", format="%d", width="small"
                    ),
                    "LAST PAGE": st.column_config.NumberColumn(
                        "LAST PAGE", format="%d", width="small"
                    ),
                    "PARTS FOUND": st.column_config.NumberColumn(
                        "PARTS FOUND", format="%d", width="small"
                    ),
                    "CONFIDENCE": st.column_config.ProgressColumn(
                        "CONFIDENCE", min_value=0, max_value=1, format="%.0f%%"
                    ),
                    "VARIANTS": st.column_config.TextColumn(
                        "VARIANTS", width="large"
                    ),
                    "DETECTION KEYS": None,
                    "ORIGIN": st.column_config.TextColumn("ORIGIN", width="small"),
                },
            )
            save_submachinery_changes = st.form_submit_button(
                "Save sub-machinery changes",
                type="primary",
                use_container_width=True,
            )

        if save_submachinery_changes:
            edited_submachineries["MCH_TP(M/S/U)"] = "SubMachinery"
            st.session_state.submachinery_review = edited_submachineries[
                SUBMACHINERY_REVIEW_COLUMNS
            ].copy()
            st.session_state.submachinery_editor_version += 1
            save_loaded_job_state()
            st.success("Sub-machinery changes saved.")
            st.rerun()

        st.caption(
            "The page range lets the reviewer open the same location in the original "
            "PDF on a second monitor. Save edits before using the action buttons above."
        )


# ---------------------------------------------------------------------------
# OCR and row extraction
# ---------------------------------------------------------------------------

with input_tab:
    active_job = active_document_job()
    if active_job:
        st.caption(f"Active document: **{active_job['file_name']}**")
    st.subheader("Step 2 — Process the scanned document")

    if main_machinery_is_ready():
        st.success(
            f"Machinery ready: {st.session_state.main_name}. You can run OCR."
        )
    else:
        st.warning(
            "Complete CODE, NAME, MAKER, and MODEL in step 1 before running OCR. "
            "This ensures extracted rows are linked to the correct machinery."
        )

    if input_type == "PDF" and source_file is not None:
        try:
            total_pages = pdf_page_count(source_file.getvalue())
            st.info(f"Uploaded PDF: **{source_file.name}** — {total_pages} pages")
            if not st.session_state.main_instruction_book:
                st.caption(
                    "Tip: enter the PDF/manual name in the Machinery tab's INSTR.BOOK field."
                )
        except Exception as exc:
            st.error(f"Could not read this PDF: {exc}")

    process_button = st.button(
        "Run OCR and extract spare-parts rows",
        type="primary",
        use_container_width=True,
        disabled=not main_machinery_is_ready(),
        help=(
            "Complete the required machinery fields in step 1 before running OCR."
            if not main_machinery_is_ready()
            else "Run OCR using the selected processing mode and active advanced settings."
        ),
    )

    if process_button:
        source_error = ""
        if not api_key:
            source_error = "A Mistral API key is required."
        elif input_type in {"PDF", "Image"} and source_file is None:
            source_error = f"Upload a {input_type.lower()} first."
        elif input_type == "Document URL" and not document_url.strip():
            source_error = "Enter a document URL first."
        elif input_type == "Image URL" and not image_url.strip():
            source_error = "Enter an image URL first."

        if source_error:
            st.error(source_error)
        else:
            progress_bar = st.progress(0.0, text="Starting OCR...")

            def show_progress(done: int, total: int, message: str) -> None:
                fraction = 0.0 if total <= 0 else min(1.0, done / total)
                progress_bar.progress(fraction, text=message)

            try:
                if input_type == "PDF":
                    pdf_bytes = source_file.getvalue()
                    total_pages = pdf_page_count(pdf_bytes)
                    selected_pages = parse_page_spec(page_spec, total_pages)
                    extracted_pages = ocr_pdf_bytes(
                        api_key=api_key,
                        pdf_bytes=pdf_bytes,
                        page_indexes=selected_pages,
                        pages_per_request=int(ocr_pages_per_request),
                        progress=show_progress,
                    )
                elif input_type == "Document URL":
                    progress_bar.progress(0.1, text="Sending document URL to OCR...")
                    extracted_pages = ocr_document_url(api_key, document_url.strip())
                elif input_type == "Image":
                    suffix = Path(source_file.name).suffix or ".png"
                    progress_bar.progress(0.1, text="Sending image to OCR...")
                    extracted_pages = ocr_image_bytes(
                        api_key,
                        source_file.getvalue(),
                        suffix,
                    )
                else:
                    progress_bar.progress(0.1, text="Sending image URL to OCR...")
                    extracted_pages = ocr_image_url(api_key, image_url.strip())

                if not extracted_pages:
                    raise RuntimeError("OCR completed but returned no pages.")

                candidate_pages, classification_frame = classify_ocr_pages(
                    extracted_pages,
                    mode=page_filter_mode,
                )
                skipped_count = len(extracted_pages) - len(candidate_pages)
                if not candidate_pages:
                    raise RuntimeError(
                        "The page filter did not find any pages to structure. "
                        "Try Conservative or Off in the sidebar."
                    )

                progress_bar.progress(
                    0.0,
                    text=(
                        f"Converting {len(candidate_pages)} candidate page(s) into "
                        "spare-parts rows..."
                    ),
                )
                extraction_messages: list[str] = []
                if structure_mode == "AI JSON extraction (recommended)":
                    rows, extraction_messages = extract_spare_parts_with_ai(
                        api_key=api_key,
                        model=extraction_model.strip() or "mistral-small-latest",
                        extracted_pages=candidate_pages,
                        pages_per_batch=int(extraction_pages_per_batch),
                        max_chars_per_batch=int(extraction_max_chars),
                        additional_instructions=extra_prompt,
                        progress=show_progress,
                    )
                    if not rows:
                        extraction_messages.append(
                            "AI extraction returned no rows; the local markdown-table "
                            "parser was used on the candidate pages."
                        )
                        rows = extract_spare_parts_from_markdown_tables(candidate_pages)
                else:
                    rows = extract_spare_parts_from_markdown_tables(candidate_pages)

                new_review = rows_to_review_dataframe(
                    rows,
                    default_machinery=st.session_state.main_name,
                    default_unit=default_unit,
                )

                if append_results:
                    merged_pages = {
                        int(page): text
                        for page, text in st.session_state.extracted_pages
                    }
                    merged_pages.update(
                        {int(page): text for page, text in extracted_pages}
                    )
                    st.session_state.extracted_pages = sorted(
                        merged_pages.items(),
                        key=lambda value: value[0],
                    )

                    previous_classification = st.session_state.page_classification
                    if previous_classification is None or previous_classification.empty:
                        combined_classification = classification_frame
                    else:
                        combined_classification = pd.concat(
                            [previous_classification, classification_frame],
                            ignore_index=True,
                        )
                        combined_classification = combined_classification.drop_duplicates(
                            subset=["SOURCE PAGE"],
                            keep="last",
                        ).sort_values("SOURCE PAGE")
                    st.session_state.page_classification = (
                        combined_classification.reset_index(drop=True)
                    )
                    st.session_state.extraction_log = list(
                        dict.fromkeys(
                            list(st.session_state.extraction_log) + extraction_messages
                        )
                    )
                    combined_review = merge_review_dataframes(
                        st.session_state.spare_review,
                        new_review,
                    )
                    previous_candidates = st.session_state.submachinery_review
                else:
                    st.session_state.extracted_pages = list(extracted_pages)
                    st.session_state.page_classification = classification_frame
                    st.session_state.extraction_log = extraction_messages
                    combined_review = new_review
                    previous_candidates = empty_submachinery_review_dataframe()

                detected_candidates = build_submachinery_candidates(
                    combined_review,
                    current_main_row(),
                )
                merged_candidates = merge_submachinery_candidates(
                    previous_candidates,
                    detected_candidates,
                )
                assigned_review = apply_submachinery_assignments(
                    combined_review,
                    merged_candidates,
                    st.session_state.main_name,
                    overwrite_auto_assignments=False,
                )

                st.session_state.submachinery_review = merged_candidates
                st.session_state.spare_review = assigned_review
                st.session_state.submachinery_editor_version += 1
                st.session_state.editor_version += 1
                st.session_state.output = None
                progress_bar.progress(1.0, text="OCR and extraction complete")
                st.success(
                    f"OCR processed {len(extracted_pages)} page(s); "
                    f"{len(candidate_pages)} were structured, {skipped_count} were skipped, "
                    f"and {len(new_review)} candidate spare-part row(s) were created. "
                    f"The app currently has {len(st.session_state.submachinery_review)} "
                    "sub-machinery proposal(s) for review."
                )
                for message in extraction_messages:
                    if message.startswith("Recovered") or "recovered" in message.lower():
                        st.info(message)
                    else:
                        st.warning(message)
            except Exception as exc:
                progress_bar.empty()
                st.error(f"Processing failed: {exc}")

    if st.session_state.extracted_pages:
        st.subheader("Raw OCR output")
        classification_lookup = {}
        if (
            st.session_state.page_classification is not None
            and not st.session_state.page_classification.empty
        ):
            classification_lookup = {
                int(row["SOURCE PAGE"]): row
                for _, row in st.session_state.page_classification.iterrows()
            }

        page_summary = pd.DataFrame(
            [
                {
                    "Page": page,
                    "Process": bool(classification_lookup.get(int(page), {}).get("PROCESS", True)),
                    "Classification": classification_lookup.get(int(page), {}).get(
                        "CLASSIFICATION", "Not classified"
                    ),
                    "Characters": len(markdown),
                    "Preview": markdown[:180].replace("\n", " "),
                }
                for page, markdown in st.session_state.extracted_pages
            ]
        )
        summary_metrics = st.columns(3)
        summary_metrics[0].metric("OCR pages", len(page_summary))
        summary_metrics[1].metric("Pages structured", int(page_summary["Process"].sum()))
        summary_metrics[2].metric(
            "Pages skipped",
            int((~page_summary["Process"]).sum()),
        )
        st.dataframe(page_summary, use_container_width=True, hide_index=True)

        raw_markdown = "\n\n".join(
            f"# Page {page}\n\n{markdown}"
            for page, markdown in st.session_state.extracted_pages
        )
        st.download_button(
            "Download raw OCR markdown",
            data=raw_markdown.encode("utf-8"),
            file_name="raw_ocr.md",
            mime="text/markdown",
        )
        with st.expander("View raw OCR markdown"):
            st.markdown(raw_markdown)

        if st.session_state.extraction_log:
            with st.expander("Extraction recovery log"):
                for message in st.session_state.extraction_log:
                    st.write(f"- {message}")


# ---------------------------------------------------------------------------
# Spare-parts review
# ---------------------------------------------------------------------------

with review_tab:
    active_job = active_document_job()
    if active_job:
        st.caption(f"Active document: **{active_job['file_name']}**")
    st.subheader("Step 4 — Review and correct candidate rows")
    st.caption(
        "Start with blocked rows, then inspect sub-machinery assignments and low-confidence "
        "identifiers. Source-page references are retained for quick navigation in the PDF."
    )
    st.info(
        f"Main machinery: **{st.session_state.main_name or '-'}**. "
        "In the table below, the **SUB-MACHINERY** column is the approved machinery "
        "record assigned to each spare part."
    )

    if st.session_state.spare_review.empty:
        st.info("Run OCR first. Candidate spare-parts rows will appear here.")
    else:
        machinery_frame = current_machinery_frame()
        valid_machinery_names = [
            str(name).strip()
            for name in machinery_frame["NAME"].tolist()
            if str(name).strip()
        ]

        full_status = recalculate_review_status(
            st.session_state.spare_review,
            valid_machinery_names=valid_machinery_names,
            allow_duplicates=False,
        )
        st.session_state.spare_review = full_status.copy()

        included_mask = full_status["INCLUDE"].astype(bool)
        ready_mask = full_status["READY"].astype(bool)
        confidence_values = pd.to_numeric(
            full_status["CONFIDENCE"], errors="coerce"
        ).fillna(0.0)
        threshold = float(st.session_state.review_confidence_threshold)
        detected_mask = full_status["DETECTED MACHINERY"].astype(str).str.strip().ne("")
        assignment_values = full_status["ASSIGNMENT SOURCE"].astype(str)
        submachinery_review_mask = (
            included_mask
            & detected_mask
            & assignment_values.eq("Main machinery default")
        )

        counts = {
            "Needs correction": int((included_mask & ~ready_mask).sum()),
            "Sub-machinery review": int(submachinery_review_mask.sum()),
            "Low confidence": int(
                (included_mask & (confidence_values < threshold)).sum()
            ),
            "Ready": int((included_mask & ready_mask).sum()),
            "Excluded": int((~included_mask).sum()),
            "All rows": len(full_status),
        }

        metric_cols = st.columns(6)
        metric_cols[0].metric("Candidates", len(full_status))
        metric_cols[1].metric("Included", int(included_mask.sum()))
        metric_cols[2].metric("Ready", counts["Ready"])
        metric_cols[3].metric("Needs correction", counts["Needs correction"])
        metric_cols[4].metric(
            "Sub-machinery review", counts["Sub-machinery review"]
        )
        metric_cols[5].metric("Low confidence", counts["Low confidence"])

        toolbar = st.columns([1.5, 1.5, 1, 1])
        with toolbar[0]:
            review_filter = st.selectbox(
                "View",
                [
                    "Needs correction",
                    "Sub-machinery review",
                    "Low confidence",
                    "Ready",
                    "Excluded",
                    "All rows",
                ],
                key="review_filter",
                format_func=lambda value: f"{value} ({counts[value]})",
                help="Needs correction is the recommended starting view.",
            )
        with toolbar[1]:
            review_sort = st.selectbox(
                "Sort by",
                [
                    "Issues first",
                    "Lowest confidence",
                    "Source page",
                    "Section start page",
                    "Sub-machinery",
                    "Part number",
                    "Description",
                ],
                key="review_sort",
            )
        with toolbar[2]:
            st.number_input(
                "Low-confidence threshold",
                min_value=0.0,
                max_value=1.0,
                step=0.05,
                format="%.2f",
                key="review_confidence_threshold",
                help="Rows below this confidence appear in Low confidence.",
            )
        with toolbar[3]:
            st.write("")
            st.write("")
            if st.button("Refresh review", use_container_width=True):
                st.session_state.editor_version += 1
                st.rerun()

        if review_filter == "Needs correction":
            visible = full_status.loc[included_mask & ~ready_mask].copy()
        elif review_filter == "Sub-machinery review":
            visible = full_status.loc[submachinery_review_mask].copy()
        elif review_filter == "Low confidence":
            visible = full_status.loc[
                included_mask & (confidence_values < threshold)
            ].copy()
        elif review_filter == "Ready":
            visible = full_status.loc[included_mask & ready_mask].copy()
        elif review_filter == "Excluded":
            visible = full_status.loc[~included_mask].copy()
        else:
            visible = full_status.copy()

        visible["_ROW_ID"] = visible.index
        if review_sort == "Issues first":
            visible["_ISSUE_RANK"] = visible["READY"].astype(bool).astype(int)
            visible = visible.sort_values(
                by=["_ISSUE_RANK", "WARNING", "CONFIDENCE", "SOURCE PAGE"],
                ascending=[True, True, True, True],
                na_position="last",
            ).drop(columns=["_ISSUE_RANK"])
        elif review_sort == "Lowest confidence":
            visible = visible.sort_values(
                "CONFIDENCE", ascending=True, na_position="last"
            )
        elif review_sort == "Source page":
            visible = visible.sort_values(
                "SOURCE PAGE", ascending=True, na_position="last"
            )
        elif review_sort == "Section start page":
            visible = visible.sort_values(
                ["SECTION START PAGE", "SOURCE PAGE"],
                ascending=True,
                na_position="last",
            )
        elif review_sort == "Sub-machinery":
            visible = visible.sort_values(
                ["MACHINERY", "SOURCE PAGE"],
                key=lambda series: series.astype(str).str.upper(),
            )
        elif review_sort == "Part number":
            visible = visible.sort_values(
                "PART NO", key=lambda series: series.astype(str).str.upper()
            )
        elif review_sort == "Description":
            visible = visible.sort_values(
                "DESCRIPTION", key=lambda series: series.astype(str).str.upper()
            )

        action_cols = st.columns([1.4, 1.4, 1.4, 1.2, 1.2])
        with action_cols[0]:
            if st.button(
                "Use main machinery for visible",
                use_container_width=True,
                disabled=visible.empty,
            ):
                frame = st.session_state.spare_review.copy()
                row_ids = visible["_ROW_ID"].tolist()
                frame.loc[row_ids, "MACHINERY"] = st.session_state.main_name
                frame.loc[row_ids, "ASSIGNMENT SOURCE"] = "Manual bulk assignment"
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()
        with action_cols[1]:
            visible_machinery = st.selectbox(
                "Set visible sub-machinery",
                [""] + valid_machinery_names,
                key="bulk_visible_machinery",
                label_visibility="collapsed",
            )
            if st.button(
                "Apply sub-machinery",
                use_container_width=True,
                disabled=visible.empty or not visible_machinery,
            ):
                frame = st.session_state.spare_review.copy()
                row_ids = visible["_ROW_ID"].tolist()
                frame.loc[row_ids, "MACHINERY"] = visible_machinery
                frame.loc[row_ids, "ASSIGNMENT SOURCE"] = "Manual bulk assignment"
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()
        with action_cols[2]:
            visible_unit = st.selectbox(
                "Set visible unit",
                ["PCS", "SET", ""],
                key="bulk_visible_unit",
                label_visibility="collapsed",
            )
            if st.button(
                "Apply unit",
                use_container_width=True,
                disabled=visible.empty,
            ):
                frame = st.session_state.spare_review.copy()
                frame.loc[visible["_ROW_ID"].tolist(), "UNIT"] = visible_unit
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()
        with action_cols[3]:
            if st.button(
                "Include visible", use_container_width=True, disabled=visible.empty
            ):
                frame = st.session_state.spare_review.copy()
                frame.loc[visible["_ROW_ID"].tolist(), "INCLUDE"] = True
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()
        with action_cols[4]:
            if st.button(
                "Exclude visible", use_container_width=True, disabled=visible.empty
            ):
                frame = st.session_state.spare_review.copy()
                frame.loc[visible["_ROW_ID"].tolist(), "INCLUDE"] = False
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()

        if visible.empty:
            if review_filter == "Needs correction":
                st.success("No included rows need correction. The review queue is clear.")
            else:
                st.info(f"No rows match the {review_filter.lower()} view.")
        else:
            st.info(f"Showing {len(visible)} of {len(full_status)} total rows.")

            source_pages = sorted(
                {
                    int(value)
                    for value in pd.to_numeric(
                        visible["SOURCE PAGE"], errors="coerce"
                    ).dropna()
                }
            )
            if source_pages:
                with st.expander("Source page quick lookup", expanded=False):
                    page_choice = st.selectbox(
                        "Source page",
                        source_pages,
                        key=(
                            f"source_page_lookup_{review_filter}_"
                            f"{st.session_state.editor_version}"
                        ),
                        help=(
                            "Open the original PDF at the same page on the second "
                            "monitor, or inspect the OCR text here."
                        ),
                    )
                    page_markdown = next(
                        (
                            markdown
                            for page, markdown in st.session_state.extracted_pages
                            if int(page) == int(page_choice)
                        ),
                        "",
                    )
                    st.caption(
                        f"Navigate to page {page_choice} in the original PDF for visual verification."
                    )
                    if page_markdown:
                        st.markdown(page_markdown)
                    else:
                        st.info("OCR text for this page is not available in the current session.")

            editor_source = visible.drop(columns=["_ROW_ID"])
            original_machinery = editor_source["MACHINERY"].astype(str).copy()
            edited_visible = st.data_editor(
                editor_source,
                key=(
                    f"spare_editor_{st.session_state.editor_version}_"
                    f"{review_filter}_{review_sort}"
                ),
                num_rows="fixed",
                use_container_width=True,
                hide_index=True,
                height=620,
                disabled=[
                    "READY",
                    "SOURCE PAGE",
                    "SECTION START PAGE",
                    "TABLE TITLE",
                    "CONFIDENCE",
                    "DETECTED MACHINERY",
                    "ASSIGNMENT SOURCE",
                    "WARNING",
                ],
                column_config={
                    "INCLUDE": st.column_config.CheckboxColumn(
                        "INCLUDE", default=True
                    ),
                    "READY": st.column_config.CheckboxColumn(
                        "READY", disabled=True
                    ),
                    "MACHINERY": st.column_config.SelectboxColumn(
                        "SUB-MACHINERY",
                        options=valid_machinery_names,
                        help=(
                            "Approved sub-machinery assigned to this spare-part row. "
                            "The main machinery appears only when no sub-machinery applies."
                        ),
                        width="medium",
                    ),
                    "PART NO": st.column_config.TextColumn(
                        "PART NO", width="medium"
                    ),
                    "DESCRIPTION": st.column_config.TextColumn(
                        "DESCRIPTION", width="large"
                    ),
                    "CODE": st.column_config.TextColumn("CODE", width="medium"),
                    "ITEM NO": st.column_config.TextColumn(
                        "ITEM NO", width="small"
                    ),
                    "UNIT": st.column_config.SelectboxColumn(
                        "UNIT", options=UNIT_OPTIONS
                    ),
                    "QNT": st.column_config.NumberColumn(
                        "QNT", min_value=0, step=1
                    ),
                    "SOURCE PAGE": st.column_config.NumberColumn(
                        "SOURCE PAGE", format="%d", width="small"
                    ),
                    "SECTION START PAGE": st.column_config.NumberColumn(
                        "SECTION START PAGE", format="%d", width="small"
                    ),
                    "TABLE TITLE": st.column_config.TextColumn(
                        "TABLE TITLE", width="large"
                    ),
                    "CONFIDENCE": st.column_config.ProgressColumn(
                        "CONFIDENCE", min_value=0, max_value=1, format="%.0f%%"
                    ),
                    "DETECTED MACHINERY": st.column_config.TextColumn(
                        "DETECTED MACHINERY", width="medium"
                    ),
                    "ASSIGNMENT SOURCE": st.column_config.TextColumn(
                        "ASSIGNMENT SOURCE", width="medium"
                    ),
                    "WARNING": st.column_config.TextColumn(
                        "WARNING", width="large"
                    ),
                },
            )

            machinery_changed = (
                edited_visible["MACHINERY"].astype(str).reset_index(drop=True)
                != original_machinery.reset_index(drop=True)
            )
            edited_visible.loc[
                machinery_changed.to_numpy(), "ASSIGNMENT SOURCE"
            ] = "Manual assignment"

            updated_full = st.session_state.spare_review.copy()
            updated_full.loc[
                visible["_ROW_ID"].tolist(), REVIEW_COLUMNS
            ] = edited_visible[REVIEW_COLUMNS].to_numpy()
            st.session_state.spare_review = recalculate_review_status(
                updated_full,
                valid_machinery_names=valid_machinery_names,
                allow_duplicates=False,
            )

            st.caption(
                "Corrections are saved immediately. Select Refresh review after edits "
                "to remove resolved rows from the current queue."
            )


# ---------------------------------------------------------------------------
# Template export, audit, and vessel email
# ---------------------------------------------------------------------------


def export_filename(vessels: list[str], machinery_code: str, machinery_name: str) -> str:
    machinery_part = safe_filename(machinery_code or machinery_name or "spare_parts")
    if len(vessels) == 1:
        vessel_part = safe_filename(vessels[0])
    elif len(vessels) == 2:
        vessel_part = "_".join(safe_filename(value) for value in vessels)
    else:
        vessel_part = f"{len(vessels)}_vessels"
    return f"{vessel_part}_{machinery_part}_import.xlsx"


def build_email_content(
    vessels: list[str],
    machinery_frame: pd.DataFrame,
    ready_rows: int,
) -> tuple[str, str]:
    vessel_subject = vessels[0] if len(vessels) == 1 else f"{len(vessels)} vessels"
    machinery_name = st.session_state.main_name or st.session_state.main_code
    subject = f"Spare Parts Import - {machinery_name} - {vessel_subject}"
    vessel_lines = "\n".join(f"- {vessel}" for vessel in vessels)
    sub_count = max(0, len(machinery_frame) - 1)
    body = f"""Dear Support Team,

Please find attached the spare-parts import workbook applicable to the following vessel(s):

{vessel_lines}

Main machinery: {st.session_state.main_name}
Code: {st.session_state.main_code}
Maker: {st.session_state.main_maker}
Model: {st.session_state.main_model}
Type: {st.session_state.main_type or '-'}
Instruction book: {st.session_state.main_instruction_book or '-'}
Included sub-machineries: {sub_count}
Ready spare-part rows: {ready_rows}

Please proceed with the corresponding import and let us know if any correction is required.

Best regards,
"""
    return subject, body



def build_multi_document_package(
    template_bytes: bytes,
    clear_existing: bool,
    allow_duplicates: bool,
) -> tuple[bytes | None, list[dict[str, str]]]:
    save_loaded_job_state()
    package = io.BytesIO()
    report: list[dict[str, str]] = []
    manifest_rows: list[dict[str, str | int]] = []
    email_sections: list[str] = []
    created_count = 0

    with zipfile.ZipFile(package, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        for position, (job_id, job) in enumerate(st.session_state.document_jobs.items(), start=1):
            vessels = _job_vessels(job)
            machinery_frame = _job_machinery_frame(job)
            machinery_errors = validate_machinery_dataframe(machinery_frame)
            valid_names = machinery_frame["NAME"].tolist() if not machinery_frame.empty else []
            review = recalculate_review_status(
                job.get("spare_review", empty_review_dataframe()),
                valid_machinery_names=valid_names,
                allow_duplicates=allow_duplicates,
            )
            included = review[review["INCLUDE"].astype(bool)] if not review.empty else review
            blocked = included[~included["READY"].astype(bool)] if not included.empty else included

            reasons = []
            if not vessels:
                reasons.append("no vessels assigned")
            reasons.extend(machinery_errors)
            if included.empty:
                reasons.append("no included spare-part rows")
            elif not blocked.empty:
                reasons.append(f"{len(blocked)} row(s) still need correction")

            if reasons:
                report.append(
                    {
                        "Document": job.get("file_name", ""),
                        "Result": "Skipped",
                        "Details": "; ".join(reasons),
                    }
                )
                continue

            try:
                workbook_bytes = build_workbook(
                    template_bytes=template_bytes,
                    machinery_frame=machinery_frame,
                    review_frame=review,
                    clear_existing=clear_existing,
                )
                file_name = export_filename(
                    vessels,
                    str(job.get("main_code", "")),
                    str(job.get("main_name", "")),
                )
                archive.writestr(f"imports/{position:02d}_{file_name}", workbook_bytes)

                audit_bytes = build_audit_workbook(
                    job.get("extracted_pages", []),
                    machinery_frame,
                    review,
                    page_classification=job.get("page_classification", pd.DataFrame()),
                    extraction_log=job.get("extraction_log", []),
                    vessels=vessels,
                    submachinery_review=job.get("submachinery_review", empty_submachinery_review_dataframe()),
                    job_metadata={
                        "Source document": job.get("file_name", ""),
                        "Main machinery code": job.get("main_code", ""),
                        "Main machinery name": job.get("main_name", ""),
                        "Maker": job.get("main_maker", ""),
                        "Model": job.get("main_model", ""),
                        "OCR pages": len(job.get("extracted_pages", [])),
                        "Included spare parts": len(included),
                    },
                )
                audit_name = safe_filename(Path(job.get("file_name", "document")).stem) + "_OCR_audit.xlsx"
                archive.writestr(f"audit/{position:02d}_{audit_name}", audit_bytes)

                vessel_lines = "\n".join(f"- {value}" for value in vessels)
                email_sections.append(
                    f"{position}. {file_name}\n"
                    f"Source PDF: {job.get('file_name', '')}\n"
                    f"Main machinery: {job.get('main_name', '')}\n"
                    f"Applicable vessels:\n{vessel_lines}\n"
                )
                manifest_rows.append(
                    {
                        "Document": job.get("file_name", ""),
                        "Import workbook": file_name,
                        "Vessels": "; ".join(vessels),
                        "Main machinery": job.get("main_name", ""),
                        "Maker": job.get("main_maker", ""),
                        "Model": job.get("main_model", ""),
                        "Sub-machineries": max(0, len(machinery_frame) - 1),
                        "Spare-part rows": len(included),
                    }
                )
                report.append(
                    {
                        "Document": job.get("file_name", ""),
                        "Result": "Included",
                        "Details": file_name,
                    }
                )
                created_count += 1
            except Exception as exc:
                report.append(
                    {
                        "Document": job.get("file_name", ""),
                        "Result": "Error",
                        "Details": str(exc),
                    }
                )

        if created_count:
            manifest = pd.DataFrame(manifest_rows)
            archive.writestr("document_vessel_assignments.csv", manifest.to_csv(index=False).encode("utf-8-sig"))
            email_body = (
                "Dear Support Team,\n\n"
                "Please find attached the import workbooks listed below. Each workbook "
                "applies only to the corresponding vessel(s).\n\n"
                + "\n".join(email_sections)
                + "\nPlease proceed with the corresponding imports and let us know if any correction is required.\n\n"
                "Best regards,\n"
            )
            archive.writestr("email_draft.txt", email_body.encode("utf-8"))

    if created_count == 0:
        return None, report
    package.seek(0)
    return package.getvalue(), report


with export_tab:
    active_job = active_document_job()
    if active_job:
        st.caption(f"Active document: **{active_job['file_name']}**")
    st.subheader("Step 5 — Build import workbook")
    vessels = selected_vessel_names()
    if vessels:
        st.info(
            "Applicable vessel(s): " + ", ".join(vessels) +
            ". Vessel names are used in the email and audit file, not in the import template."
        )
    else:
        st.error("Select at least one vessel in step 1 before export.")

    st.markdown(
        "The app writes the reviewed data into the existing template as follows:\n\n"
        "- **1.Machineries|Sub|Units**, starting at row 5: CODE, NAME, MAKER, MODEL, "
        "TYPE, INSTR.BOOK, SPECIFICATIONS, MCH_TP(M/S/U).\n"
        "- **2.Spare Parts**, starting at row 4: MACHINERY (approved sub-machinery), PART NO, DESCRIPTION, CODE, "
        "ITEM NO, UNIT, QNT.\n\n"
        "Source pages, detected headings, confidence values, vessels, and review notes "
        "remain in the separate audit workbook."
    )

    clear_existing = st.checkbox(
        "Clear existing data rows in the template before writing",
        value=True,
        help=(
            "Recommended for a new import file. Turn off only when intentionally "
            "adding to records already stored in a replacement template."
        ),
    )
    allow_duplicates = st.checkbox(
        "Allow possible duplicate spare-part rows",
        value=False,
        help="Keep this off unless repeated rows are intentional.",
    )

    machinery_frame = current_machinery_frame()
    machinery_errors = validate_machinery_dataframe(machinery_frame)
    valid_machinery_names = machinery_frame["NAME"].tolist()
    export_review = recalculate_review_status(
        st.session_state.spare_review,
        valid_machinery_names=valid_machinery_names,
        allow_duplicates=allow_duplicates,
    )
    st.session_state.spare_review = export_review

    included = export_review[export_review["INCLUDE"].astype(bool)]
    blocked = included[~included["READY"].astype(bool)]

    if machinery_errors:
        for error in machinery_errors:
            st.error(error)
    if blocked.empty and not included.empty:
        st.success(f"{len(included)} included spare-part row(s) are ready.")
    elif included.empty:
        st.warning("No spare-part rows are currently included.")
    else:
        st.error(f"Correct {len(blocked)} included row(s) before export.")
        st.dataframe(
            blocked[
                [
                    "SOURCE PAGE",
                    "SECTION START PAGE",
                    "MACHINERY",
                    "PART NO",
                    "DESCRIPTION",
                    "ITEM NO",
                    "WARNING",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

    can_build = (
        template_bytes is not None
        and bool(vessels)
        and not machinery_errors
        and not included.empty
        and blocked.empty
    )

    if st.button(
        "Create Excel",
        type="primary",
        disabled=not can_build,
        use_container_width=True,
    ):
        try:
            output_bytes = build_workbook(
                template_bytes=template_bytes,
                machinery_frame=machinery_frame,
                review_frame=export_review,
                clear_existing=clear_existing,
            )
            st.session_state.output_name = export_filename(
                vessels,
                st.session_state.main_code,
                st.session_state.main_name,
            )
            st.session_state.output = output_bytes
            email_subject, email_body = build_email_content(
                vessels,
                machinery_frame,
                len(included),
            )
            st.session_state.prepared_email_subject = email_subject
            st.session_state.prepared_email_body = email_body
            st.success(
                "Workbook created. Download it below and test a small import first."
            )
        except Exception as exc:
            st.error(f"Could not create the workbook: {exc}")

    if st.session_state.output:
        st.download_button(
            "Download import workbook",
            data=st.session_state.output,
            file_name=st.session_state.output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        st.subheader("Prepared email")
        if not st.session_state.prepared_email_subject or not st.session_state.prepared_email_body:
            email_subject, email_body = build_email_content(
                vessels,
                machinery_frame,
                len(included),
            )
            st.session_state.prepared_email_subject = email_subject
            st.session_state.prepared_email_body = email_body
        email_subject_value = st.text_input(
            "Subject",
            key="prepared_email_subject",
        )
        email_body_value = st.text_area(
            "Body",
            height=330,
            key="prepared_email_body",
            help="Review, copy, and paste this text into the email accompanying the workbook.",
        )
        email_text = f"Subject: {email_subject_value}\n\n{email_body_value}"
        st.download_button(
            "Download email draft",
            data=email_text.encode("utf-8"),
            file_name=safe_filename(st.session_state.output_name) + "_email.txt",
            mime="text/plain",
            use_container_width=True,
        )

    if len(st.session_state.document_jobs) > 1:
        st.divider()
        st.subheader("All documents package")
        st.caption(
            "Builds one import workbook and one audit workbook for every ready PDF. "
            "A single PDF assigned to several vessels is included only once, with the vessel mapping in the manifest and email draft."
        )
        if st.button(
            "Build ZIP package for all ready documents",
            use_container_width=True,
            disabled=template_bytes is None,
        ):
            package_bytes, package_report = build_multi_document_package(
                template_bytes=template_bytes,
                clear_existing=clear_existing,
                allow_duplicates=allow_duplicates,
            )
            st.session_state.multi_package_output = package_bytes
            st.session_state.multi_package_report = package_report
            if package_bytes:
                st.success("Multi-document package created.")
            else:
                st.error("No document was ready for package export. Review the report below.")

        if st.session_state.multi_package_report:
            st.dataframe(
                pd.DataFrame(st.session_state.multi_package_report),
                use_container_width=True,
                hide_index=True,
            )
        if st.session_state.multi_package_output:
            st.download_button(
                "Download all-documents ZIP package",
                data=st.session_state.multi_package_output,
                file_name=st.session_state.multi_package_name,
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )

    if st.session_state.extracted_pages:
        audit_bytes = build_audit_workbook(
            st.session_state.extracted_pages,
            machinery_frame,
            export_review,
            page_classification=st.session_state.page_classification,
            extraction_log=st.session_state.extraction_log,
            vessels=vessels,
            submachinery_review=st.session_state.submachinery_review,
            job_metadata={
                "Main machinery code": st.session_state.main_code,
                "Main machinery name": st.session_state.main_name,
                "Maker": st.session_state.main_maker,
                "Model": st.session_state.main_model,
                "Type": st.session_state.main_type,
                "Instruction book": st.session_state.main_instruction_book,
                "Processing mode": st.session_state.processing_preset,
                "OCR pages": len(st.session_state.extracted_pages),
                "Included spare parts": len(included),
            },
        )
        audit_name = (
            safe_filename(st.session_state.output_name or st.session_state.main_name)
            + "_OCR_audit.xlsx"
        )
        st.download_button(
            "Download OCR audit/review workbook",
            data=audit_bytes,
            file_name=audit_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )



# Persist the active document workspace after every completed rerun.
save_loaded_job_state()
