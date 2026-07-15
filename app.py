from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import streamlit as st

from tools import (
    MACHINERY_COLUMNS,
    MACHINERY_TYPES,
    PAGE_FILTER_MODES,
    REVIEW_COLUMNS,
    UNIT_OPTIONS,
    build_audit_workbook,
    build_workbook,
    classify_ocr_pages,
    empty_additional_machinery_dataframe,
    empty_review_dataframe,
    extract_spare_parts_from_markdown_tables,
    extract_spare_parts_with_ai,
    machinery_rows_from_main_and_additional,
    merge_review_dataframes,
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
APP_VERSION = "2.4"

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
        "additional_machinery": empty_additional_machinery_dataframe(),
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
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def get_secret(name: str) -> str:
    try:
        return str(st.secrets[name])
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

st.title("📄 Spare Parts OCR Import Builder")


# ---------------------------------------------------------------------------
# Sidebar settings
# ---------------------------------------------------------------------------

with st.sidebar:
    with st.expander("📖 Instructions & Help", expanded=False):
        st.markdown(
            """
### Quick start

1. Open **1. Machinery** and complete `CODE`, `NAME`, `MAKER`, and `MODEL`.
2. Under **Source**, upload the scanned PDF and choose the pages to process.
3. Select a processing mode. **Balanced** is the recommended default.
4. Open **2. OCR** and run the extraction.
5. Open **3. Review spare parts**, correct warnings, and exclude unwanted rows.
6. Open **4. Export**, create the workbook, and test a small import first.

### Processing modes

**Balanced — recommended**  
Best for most manuals. Uses moderate OCR and AI batches for a good balance of speed and stability.

**Fast**  
Use for clean, consistent scans and regular tables. It processes larger batches and may require more review.

**Careful**  
Use for poor scans, complex tables, missing OCR pages, or repeated JSON recovery messages. It processes smaller batches and is slower but more stable.

### Advanced Mistral settings

Normal users only need to select a processing mode. Open **Advanced Mistral settings** when a manual needs fine-tuning. The expander shows the exact active values for:

- page filtering,
- extraction method and model,
- PDF pages per OCR request,
- OCR pages per structuring batch,
- maximum characters per AI batch,
- default unit, and
- manual-specific extraction instructions.

Selecting a processing mode resets the advanced numeric settings to that mode's defaults. Advanced changes then apply to the current run.

### Source and page ranges

**Pages to process**  
Use `all`, `1-20`, `25`, or `30-45`. For large books, process ranges and enable **Append to current review table** after the first range.

### Review dashboard

The Review tab opens on **Needs correction** by default, so users see only blocked rows first.

- Use **View** to switch between Needs correction, Low confidence, Ready, Excluded, and All rows.
- Use **Sort by** for issues first, lowest confidence, source page, part number, or description.
- Correct values directly in the visible table. Changes are written back to the full review dataset automatically.
- Use **Set visible unit** and **Set visible machinery** for quick bulk corrections.
- **INCLUDE:** checked rows are considered for export.
- **READY:** calculated automatically after validation.
- **WARNING:** explains what blocks export.
- **SOURCE PAGE:** audit reference; it is not written to the import template.

### Troubleshooting

- **No candidate pages:** in Advanced settings, change page filtering to Conservative or Off.
- **Repeated JSON recovery messages:** choose **Careful**, or reduce structuring batch and maximum characters in Advanced settings.
- **Missing OCR pages or timeouts:** choose **Careful**, or reduce PDF pages per OCR request.
- **Too many false spare parts:** use Strict page filtering and add a short manual-specific instruction.
- **Part numbers changed:** add an instruction to preserve leading zeros, spaces, slashes, dots, and hyphens exactly.
- **Large job:** process page ranges and download the audit workbook regularly. A redeploy or session reset clears in-memory results.

### Data handling

Uploaded pages are sent to the configured Mistral service when OCR or AI extraction runs. Use the tool only for documents approved for that processing.
            """
        )

    st.header("Source")
    input_type = st.radio(
        "Choose input type",
        ["PDF", "Document URL", "Image", "Image URL"],
        index=0,  # PDF is deliberately the default.
        help="PDF is recommended for scanned manuals. URL options require an accessible direct file URL.",
    )

    source_file = None
    document_url = ""
    image_url = ""
    page_spec = ""

    if input_type == "PDF":
        source_file = st.file_uploader(
            "Upload a scanned PDF",
            type=["pdf"],
            help="Upload the original scanned manual. For very large books, process selected page ranges.",
        )
        page_spec = st.text_input(
            "Pages to process",
            value="all",
            help="Examples: all, 1-20, 25, 30-35. Process large books in batches.",
        )
    elif input_type == "Document URL":
        document_url = st.text_input(
            "Document URL",
            help="Enter a direct, publicly accessible PDF URL. Use file upload for internal documents.",
        )
    elif input_type == "Image":
        source_file = st.file_uploader(
            "Upload an image",
            type=["png", "jpg", "jpeg"],
            help="Use this for a single scanned page or photograph.",
        )
    else:
        image_url = st.text_input(
            "Image URL",
            help="Enter a direct, publicly accessible image URL.",
        )

    append_results = st.checkbox(
        "Append to current review table",
        value=False,
        help="Useful when processing different page ranges from the same large manual.",
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
        "Reset OCR and review data",
        use_container_width=True,
        help="Clears extracted pages, classifications, candidate rows and the generated workbook from this session.",
    ):
        st.session_state.extracted_pages = []
        st.session_state.page_classification = pd.DataFrame()
        st.session_state.extraction_log = []
        st.session_state.spare_review = empty_review_dataframe()
        st.session_state.output = None
        st.session_state.editor_version += 1
        st.rerun()

    st.divider()
    with st.expander("ℹ️ About", expanded=False):
        st.markdown(
            f"""
**Spare Parts OCR Import Builder — v{APP_VERSION}**

**Workflow**  
Machinery → Upload → OCR → Page filtering → AI extraction → Review → Excel export

**Supported sources**  
Scanned PDFs, document URLs, images and image URLs.

**Important**  
The generated workbook should be tested with a small import batch before production use.
            """
        )


# ---------------------------------------------------------------------------
# Main machinery and optional sub-units
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


def main_machinery_is_ready() -> bool:
    required_keys = ("main_code", "main_name", "main_maker", "main_model")
    return all(str(st.session_state.get(key, "")).strip() for key in required_keys)


machinery_tab, input_tab, review_tab, export_tab = st.tabs(
    ["1. Machinery", "2. OCR", "3. Review spare parts", "4. Export"]
)

with machinery_tab:
    st.subheader("Step 1 — Main machinery")
    st.info(
        "Requires CODE, NAME, MAKER, MODEL and MCH_TP. The app fixes the "
        "main row's MCH_TP to 'Main Machinery'."
    )

    row1 = st.columns(4)
    with row1[0]:
        st.text_input("CODE *", key="main_code", placeholder="M/E", help="Required machinery code written to sheet 1.")
    with row1[1]:
        st.text_input("NAME *", key="main_name", placeholder="Main Engine", help="Required. Spare-part MACHINERY values must match this name exactly.")
    with row1[2]:
        st.text_input("MAKER *", key="main_maker", placeholder="Sulzer", help="Required machinery manufacturer.")
    with row1[3]:
        st.text_input("MODEL *", key="main_model", placeholder="RTA 56", help="Required machinery model.")

    row2 = st.columns(3)
    with row2[0]:
        st.text_input("TYPE", key="main_type", help="Optional type or variant from the manual/nameplate.")
    with row2[1]:
        st.text_input("INSTR.BOOK", key="main_instruction_book", help="Instruction-book/manual reference. The uploaded PDF filename is filled automatically when empty.")
    with row2[2]:
        st.text_input("SPECIFICATIONS", key="main_specifications", help="Optional technical specifications or distinguishing information.")

    st.text_input(
        "MCH_TP(M/S/U)",
        value="Main Machinery",
        disabled=True,
        key="fixed_main_machinery_type",
    )

    st.subheader("Optional sub-machineries")
    st.caption(
        "Add rows only when spare parts must refer to a sub-machinery instead "
        "of the main machinery. Every added row is automatically classified as SubMachinery."
    )

    # Normalize any existing rows from older app versions.
    additional_frame = st.session_state.additional_machinery.copy()
    if not additional_frame.empty:
        additional_frame["MCH_TP(M/S/U)"] = "SubMachinery"
        st.session_state.additional_machinery = additional_frame

    button_cols = st.columns([1, 1, 3])
    with button_cols[0]:
        if st.button("Add sub-machinery", use_container_width=True):
            new_row = {column: "" for column in MACHINERY_COLUMNS}
            new_row["MCH_TP(M/S/U)"] = "SubMachinery"
            st.session_state.additional_machinery = pd.concat(
                [
                    st.session_state.additional_machinery,
                    pd.DataFrame([new_row], columns=MACHINERY_COLUMNS),
                ],
                ignore_index=True,
            )
            st.session_state.submachinery_editor_version = (
                st.session_state.get("submachinery_editor_version", 0) + 1
            )
            st.rerun()

    with button_cols[1]:
        if st.button("Clear sub-machineries", use_container_width=True):
            st.session_state.additional_machinery = empty_additional_machinery_dataframe()
            st.session_state.submachinery_editor_version = (
                st.session_state.get("submachinery_editor_version", 0) + 1
            )
            st.rerun()

    if st.session_state.additional_machinery.empty:
        st.info("No sub-machineries added yet. Select **Add sub-machinery** to create a row.")
    else:
        additional_editor = st.data_editor(
            st.session_state.additional_machinery,
            key=f"additional_machinery_editor_{st.session_state.get('submachinery_editor_version', 0)}",
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            disabled=["MCH_TP(M/S/U)"],
            column_config={
                "CODE": st.column_config.TextColumn("CODE", help="Sub-machinery code."),
                "NAME": st.column_config.TextColumn("NAME", help="Sub-machinery name used by spare-part rows."),
                "MAKER": st.column_config.TextColumn("MAKER"),
                "MODEL": st.column_config.TextColumn("MODEL"),
                "TYPE": st.column_config.TextColumn("TYPE"),
                "INSTR.BOOK": st.column_config.TextColumn("INSTR.BOOK"),
                "SPECIFICATIONS": st.column_config.TextColumn("SPECIFICATIONS"),
                "MCH_TP(M/S/U)": st.column_config.TextColumn(
                    "MCH_TP(M/S/U)",
                    help="Automatically fixed to SubMachinery.",
                ),
            },
        )
        additional_editor["MCH_TP(M/S/U)"] = "SubMachinery"
        st.session_state.additional_machinery = additional_editor.copy()


# ---------------------------------------------------------------------------
# OCR and row extraction
# ---------------------------------------------------------------------------

with input_tab:
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
                    st.session_state.page_classification = combined_classification.reset_index(drop=True)
                    st.session_state.extraction_log = list(
                        dict.fromkeys(
                            list(st.session_state.extraction_log) + extraction_messages
                        )
                    )
                    st.session_state.spare_review = merge_review_dataframes(
                        st.session_state.spare_review,
                        new_review,
                    )
                else:
                    st.session_state.extracted_pages = list(extracted_pages)
                    st.session_state.page_classification = classification_frame
                    st.session_state.extraction_log = extraction_messages
                    st.session_state.spare_review = new_review

                st.session_state.editor_version += 1
                st.session_state.output = None
                progress_bar.progress(1.0, text="OCR and extraction complete")
                st.success(
                    f"OCR processed {len(extracted_pages)} page(s); "
                    f"{len(candidate_pages)} were structured, {skipped_count} were skipped, "
                    f"and {len(new_review)} candidate spare-part row(s) were created."
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


def current_machinery_frame() -> pd.DataFrame:
    return machinery_rows_from_main_and_additional(
        current_main_row(),
        st.session_state.additional_machinery,
    )


with review_tab:
    st.subheader("Step 3 — Review and correct candidate rows")
    st.caption(
        "The table opens on rows that need correction. Filter and sort instantly, "
        "edit the visible rows, and the app writes those changes back to the complete dataset."
    )

    if st.session_state.spare_review.empty:
        st.info("Run OCR first. Candidate spare-parts rows will appear here.")
    else:
        machinery_frame = current_machinery_frame()
        valid_machinery_names = machinery_frame["NAME"].tolist()

        full_status = recalculate_review_status(
            st.session_state.spare_review,
            valid_machinery_names=valid_machinery_names,
            allow_duplicates=False,
        )
        st.session_state.spare_review = full_status.copy()

        included_mask = full_status["INCLUDE"].astype(bool)
        ready_mask = full_status["READY"].astype(bool)
        confidence_values = pd.to_numeric(full_status["CONFIDENCE"], errors="coerce").fillna(0.0)
        threshold = float(st.session_state.review_confidence_threshold)

        counts = {
            "Needs correction": int((included_mask & ~ready_mask).sum()),
            "Low confidence": int((included_mask & (confidence_values < threshold)).sum()),
            "Ready": int((included_mask & ready_mask).sum()),
            "Excluded": int((~included_mask).sum()),
            "All rows": len(full_status),
        }

        metric_cols = st.columns(5)
        metric_cols[0].metric("Candidates", len(full_status))
        metric_cols[1].metric("Included", int(included_mask.sum()))
        metric_cols[2].metric("Ready", counts["Ready"])
        metric_cols[3].metric("Needs correction", counts["Needs correction"])
        metric_cols[4].metric("Low confidence", counts["Low confidence"])

        toolbar = st.columns([1.5, 1.5, 1, 1])
        with toolbar[0]:
            review_filter = st.selectbox(
                "View",
                ["Needs correction", "Low confidence", "Ready", "Excluded", "All rows"],
                key="review_filter",
                format_func=lambda value: f"{value} ({counts[value]})",
                help="Needs correction is the recommended default for fast quality control.",
            )
        with toolbar[1]:
            review_sort = st.selectbox(
                "Sort by",
                ["Issues first", "Lowest confidence", "Source page", "Part number", "Description"],
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
                help="Rows below this confidence appear in the Low confidence view.",
            )
        with toolbar[3]:
            st.write("")
            st.write("")
            if st.button("Refresh review", use_container_width=True):
                st.session_state.editor_version += 1
                st.rerun()

        if review_filter == "Needs correction":
            visible = full_status.loc[included_mask & ~ready_mask].copy()
        elif review_filter == "Low confidence":
            visible = full_status.loc[included_mask & (confidence_values < threshold)].copy()
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
            visible = visible.sort_values("CONFIDENCE", ascending=True, na_position="last")
        elif review_sort == "Source page":
            visible = visible.sort_values("SOURCE PAGE", ascending=True, na_position="last")
        elif review_sort == "Part number":
            visible = visible.sort_values("PART NO", key=lambda s: s.astype(str).str.upper())
        elif review_sort == "Description":
            visible = visible.sort_values("DESCRIPTION", key=lambda s: s.astype(str).str.upper())

        action_cols = st.columns([1.4, 1.4, 1.4, 1.2, 1.2])
        with action_cols[0]:
            if st.button("Use main machinery for visible", use_container_width=True, disabled=visible.empty):
                frame = st.session_state.spare_review.copy()
                frame.loc[visible["_ROW_ID"], "MACHINERY"] = st.session_state.main_name
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()
        with action_cols[1]:
            visible_machinery = st.selectbox(
                "Set visible machinery",
                [""] + [name for name in valid_machinery_names if str(name).strip()],
                key="bulk_visible_machinery",
                label_visibility="collapsed",
            )
            if st.button("Apply machinery", use_container_width=True, disabled=visible.empty or not visible_machinery):
                frame = st.session_state.spare_review.copy()
                frame.loc[visible["_ROW_ID"], "MACHINERY"] = visible_machinery
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
            if st.button("Apply unit", use_container_width=True, disabled=visible.empty):
                frame = st.session_state.spare_review.copy()
                frame.loc[visible["_ROW_ID"], "UNIT"] = visible_unit
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()
        with action_cols[3]:
            if st.button("Include visible", use_container_width=True, disabled=visible.empty):
                frame = st.session_state.spare_review.copy()
                frame.loc[visible["_ROW_ID"], "INCLUDE"] = True
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()
        with action_cols[4]:
            if st.button("Exclude visible", use_container_width=True, disabled=visible.empty):
                frame = st.session_state.spare_review.copy()
                frame.loc[visible["_ROW_ID"], "INCLUDE"] = False
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
            editor_source = visible.drop(columns=["_ROW_ID"])
            edited_visible = st.data_editor(
                editor_source,
                key=f"spare_editor_{st.session_state.editor_version}_{review_filter}_{review_sort}",
                num_rows="fixed",
                use_container_width=True,
                hide_index=True,
                height=620,
                disabled=[
                    "READY",
                    "SOURCE PAGE",
                    "CONFIDENCE",
                    "DETECTED MACHINERY",
                    "WARNING",
                ],
                column_config={
                    "INCLUDE": st.column_config.CheckboxColumn("INCLUDE", default=True),
                    "READY": st.column_config.CheckboxColumn("READY", disabled=True),
                    "MACHINERY": st.column_config.SelectboxColumn(
                        "MACHINERY",
                        options=[name for name in valid_machinery_names if str(name).strip()],
                        help="Must match a machinery name entered in step 1.",
                        width="medium",
                    ),
                    "PART NO": st.column_config.TextColumn("PART NO", width="medium"),
                    "DESCRIPTION": st.column_config.TextColumn("DESCRIPTION", width="large"),
                    "CODE": st.column_config.TextColumn("CODE", width="medium"),
                    "ITEM NO": st.column_config.TextColumn("ITEM NO", width="small"),
                    "UNIT": st.column_config.SelectboxColumn("UNIT", options=UNIT_OPTIONS),
                    "QNT": st.column_config.NumberColumn("QNT", min_value=0, step=1),
                    "SOURCE PAGE": st.column_config.NumberColumn("SOURCE PAGE", format="%d"),
                    "CONFIDENCE": st.column_config.ProgressColumn(
                        "CONFIDENCE", min_value=0, max_value=1, format="%.0f%%"
                    ),
                    "DETECTED MACHINERY": st.column_config.TextColumn(
                        "DETECTED MACHINERY", width="medium"
                    ),
                    "WARNING": st.column_config.TextColumn("WARNING", width="large"),
                },
            )

            updated_full = st.session_state.spare_review.copy()
            updated_full.loc[visible["_ROW_ID"], REVIEW_COLUMNS] = edited_visible[REVIEW_COLUMNS].to_numpy()
            st.session_state.spare_review = recalculate_review_status(
                updated_full,
                valid_machinery_names=valid_machinery_names,
                allow_duplicates=False,
            )

            st.caption(
                "Corrections are saved immediately. After editing, select Refresh review "
                "to remove rows that are now ready from the Needs correction view."
            )


# ---------------------------------------------------------------------------
# Template export
# ---------------------------------------------------------------------------

with export_tab:
    st.subheader("Build import workbook")
    st.markdown(
        "The app writes the reviewed data into the existing template as follows:\n\n"
        "- **1.Machineries|Sub|Units**, starting at row 5: CODE, NAME, MAKER, MODEL, "
        "TYPE, INSTR.BOOK, SPECIFICATIONS, MCH_TP(M/S/U).\n"
        "- **2.Spare Parts**, starting at row 4: MACHINERY, PART NO, DESCRIPTION, CODE, "
        "ITEM NO, UNIT, QNT."
    )

    clear_existing = st.checkbox(
        "Clear existing data rows in the template before writing",
        value=True,
        help="Recommended when creating a new import file. Turn off only when intentionally adding to data already stored in the uploaded template.",
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
            blocked[[
                "SOURCE PAGE",
                "MACHINERY",
                "PART NO",
                "DESCRIPTION",
                "ITEM NO",
                "WARNING",
            ]],
            use_container_width=True,
            hide_index=True,
        )

    can_build = (
        template_bytes is not None
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
            name_source = (
                st.session_state.main_code
                or st.session_state.main_name
                or "spare_parts"
            )
            output_name = safe_filename(name_source) + "_import.xlsx"
            st.session_state.output = output_bytes
            st.session_state.output_name = output_name
            st.success("Workbook created. Download it below and test a small import first.")
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

    if st.session_state.extracted_pages:
        audit_bytes = build_audit_workbook(
            st.session_state.extracted_pages,
            machinery_frame,
            export_review,
            page_classification=st.session_state.page_classification,
            extraction_log=st.session_state.extraction_log,
        )
        st.download_button(
            "Download OCR audit/review workbook",
            data=audit_bytes,
            file_name="OCR_review_and_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
