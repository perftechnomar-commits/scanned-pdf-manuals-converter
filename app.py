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
APP_VERSION = "2.1"

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
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def get_secret(name: str) -> str:
    try:
        return str(st.secrets[name])
    except Exception:
        return ""


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

1. Open **2. Machinery** and complete the required main-machinery fields.
2. Upload the scanned manual under **Source**. Start with a small page range such as `1-20`.
3. Keep the recommended Mistral settings unless the manual is difficult.
4. Open **1. OCR** and select **Run OCR and extract spare-parts rows**.
5. Open **3. Review spare parts**, correct warnings and exclude unwanted rows.
6. Open **4. Export**, create the workbook and test a small import first.

### Recommended settings

- **Page filtering:** Conservative
- **Structured-extraction model:** `mistral-small-latest`
- **PDF pages per OCR request:** `25`
- **OCR pages per structuring batch:** `3`
- **Maximum OCR characters per AI batch:** `12000`
- **Default unit:** `PCS`

### What the settings mean

**Pages to process**  
Use `all`, `1-20`, `25`, or `30-45`. For large books, process ranges and enable **Append to current review table**.

**Page filtering**  
- **Conservative:** recommended; skips only obvious non-parts pages.
- **Strict:** use when many contents, narrative, or drawing-only pages create false rows.
- **Off:** processes every OCR page; useful if valid tables are being skipped.

**PDF pages per OCR request**  
Controls OCR upload size. Use `25` normally, `10-15` for poor scans or failed OCR requests.

**OCR pages per structuring batch**  
Controls how much OCR text is converted to JSON at once. Use `3` normally and `1-2` for difficult layouts or repeated JSON recovery messages.

**Maximum OCR characters per AI batch**  
Use `12000` normally. Reduce to `8000` if responses are truncated or malformed.

**Optional extraction instructions**  
Use only for rules specific to the current manual, for example:  
`Preserve leading zeros and hyphens. The first column is ITEM NO and the second is PART NO. Ignore prices and drawing dimensions.`

### Review-table guide

- **INCLUDE:** checked rows are considered for export.
- **READY:** calculated automatically after validation.
- **MACHINERY:** must exactly match a machinery name entered in tab 2.
- **CONFIDENCE:** low values should be checked against the source page.
- **WARNING:** explains what must be corrected before export.
- **SOURCE PAGE:** audit reference; it is not written to the import template.

### Troubleshooting

- **No candidate pages:** change page filtering to Conservative or Off.
- **Repeated JSON recovery messages:** reduce structuring batch to `1-2` and maximum characters to `8000`.
- **Missing OCR pages:** reduce PDF pages per OCR request to `10-15`.
- **Too many false spare parts:** use Strict filtering and add a short manual-specific instruction.
- **Part numbers changed:** instruct the model to preserve leading zeros, spaces, slashes and hyphens exactly.
- **Large job:** process the manual in ranges and download the audit workbook regularly. A redeploy or session reset clears in-memory results.

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
    st.header("Mistral settings")
    secret_api_key = get_secret("MISTRAL_API_KEY")
    if secret_api_key:
        st.success("MISTRAL_API_KEY loaded from Streamlit secrets.")
        entered_api_key = ""
    else:
        entered_api_key = st.text_input(
            "Mistral API key",
            type="password",
            help="For local testing only. Prefer .streamlit/secrets.toml for deployment.",
        )
    api_key = secret_api_key or entered_api_key

    structure_mode = st.selectbox(
        "Convert OCR text into rows",
        ["AI JSON extraction (recommended)", "Local markdown-table parser"],
        index=0,
        help=(
            "AI JSON extraction handles irregular tables and wrapped descriptions. "
            "The local parser is faster but works best only when OCR already produced clean Markdown tables."
        ),
    )
    page_filter_mode = st.selectbox(
        "Page filtering before AI extraction",
        PAGE_FILTER_MODES,
        index=0,
        help=(
            "Conservative skips only obvious contents/revision/prose pages. Strict "
            "processes only strong parts-table candidates. Off processes every OCR page. "
            "This filter is local and makes no extra API calls."
        ),
    )
    extraction_model = st.text_input(
        "Structured-extraction model",
        value="mistral-small-latest",
        disabled=structure_mode != "AI JSON extraction (recommended)",
        help="Recommended default: mistral-small-latest. Change only after testing another supported Mistral model.",
    )
    ocr_pages_per_request = st.number_input(
        "PDF pages per OCR request",
        min_value=1,
        max_value=100,
        value=25,
        step=1,
        help="Recommended: 25. Use 10-15 for poor scans, missing pages, timeouts, or failed OCR requests.",
    )
    extraction_pages_per_batch = st.number_input(
        "OCR pages per structuring batch",
        min_value=1,
        max_value=20,
        value=3,
        step=1,
        disabled=structure_mode != "AI JSON extraction (recommended)",
        help="Recommended: 3. Use 1-2 for complex layouts or repeated malformed/truncated JSON responses.",
    )
    extraction_max_chars = st.number_input(
        "Maximum OCR characters per AI batch",
        min_value=2000,
        max_value=30000,
        value=12000,
        step=1000,
        disabled=structure_mode != "AI JSON extraction (recommended)",
        help=(
            "Smaller batches reduce malformed/truncated JSON. If a response still "
            "fails, the app automatically divides it into smaller requests."
        ),
    )
    default_unit = st.selectbox(
        "Default spare-part unit",
        ["PCS", "SET", ""],
        index=0,
        help="PCS is the normal default. Use SET for kits/sets, or blank when every unit must be reviewed manually.",
    )
    extra_prompt = st.text_area(
        "Optional manual-specific extraction instructions",
        placeholder=(
            "Example: The first column is ITEM NO and the second column is PART NO. "
            "Ignore drawing dimensions and prices."
        ),
        height=100,
        disabled=structure_mode != "AI JSON extraction (recommended)",
        help=(
            "Add only manual-specific rules. Example: Preserve leading zeros and hyphens; "
            "the first column is ITEM NO; ignore prices and drawing dimensions."
        ),
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
Upload → OCR → Page filtering → AI extraction → Review → Excel export

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


input_tab, machinery_tab, review_tab, export_tab = st.tabs(
    ["1. OCR", "2. Machinery", "3. Review spare parts", "4. Export"]
)

with machinery_tab:
    st.subheader("Main machinery — written to row 5 of sheet 1")
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

    st.subheader("Optional sub-machineries or book units")
    st.caption(
        "Add rows only when spare parts must refer to a sub-machinery or unit instead "
        "of the main machinery. All required fields must be filled for each row."
    )

    additional_editor = st.data_editor(
        st.session_state.additional_machinery,
        key="additional_machinery_editor",
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "MCH_TP(M/S/U)": st.column_config.SelectboxColumn(
                "MCH_TP(M/S/U)",
                options=["SubMachinery", "Unit (Book Chapter)"],
            )
        },
    )
    st.session_state.additional_machinery = additional_editor


# ---------------------------------------------------------------------------
# OCR and row extraction
# ---------------------------------------------------------------------------

with input_tab:
    st.subheader("Process the scanned document")

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
    st.subheader("Review and correct the candidate rows")
    st.caption(
        "Will receive only rows where INCLUDE and READY are both checked. "
        "SOURCE PAGE, CONFIDENCE, DETECTED MACHINERY, and WARNING are audit fields and "
        "are not written to the template."
    )

    if st.session_state.spare_review.empty:
        st.info("Run OCR first. Candidate spare-parts rows will appear here.")
    else:
        machinery_frame = current_machinery_frame()
        valid_machinery_names = machinery_frame["NAME"].tolist()

        top_buttons = st.columns(3)
        with top_buttons[0]:
            if st.button("Use main machinery for all included rows"):
                frame = st.session_state.spare_review.copy()
                mask = frame["INCLUDE"].astype(bool)
                frame.loc[mask, "MACHINERY"] = st.session_state.main_name
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()
        with top_buttons[1]:
            if st.button("Select all candidate rows"):
                frame = st.session_state.spare_review.copy()
                frame["INCLUDE"] = True
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()
        with top_buttons[2]:
            if st.button("Exclude all rows"):
                frame = st.session_state.spare_review.copy()
                frame["INCLUDE"] = False
                st.session_state.spare_review = frame
                st.session_state.editor_version += 1
                st.rerun()

        status_frame = recalculate_review_status(
            st.session_state.spare_review,
            valid_machinery_names=valid_machinery_names,
            allow_duplicates=False,
        )

        edited = st.data_editor(
            status_frame,
            key=f"spare_editor_{st.session_state.editor_version}",
            num_rows="dynamic",
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
                "MACHINERY": st.column_config.TextColumn(
                    "MACHINERY",
                    help="Must exactly match a NAME entered on sheet 1.",
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
                    "CONFIDENCE",
                    min_value=0,
                    max_value=1,
                    format="%.0f%%",
                ),
                "DETECTED MACHINERY": st.column_config.TextColumn(
                    "DETECTED MACHINERY", width="medium"
                ),
                "WARNING": st.column_config.TextColumn("WARNING", width="large"),
            },
        )

        # Recalculate after edits so export always uses the latest values.
        st.session_state.spare_review = recalculate_review_status(
            edited,
            valid_machinery_names=valid_machinery_names,
            allow_duplicates=False,
        )

        frame = st.session_state.spare_review
        included_count = int(frame["INCLUDE"].astype(bool).sum())
        ready_count = int((frame["INCLUDE"].astype(bool) & frame["READY"].astype(bool)).sum())
        blocked_count = included_count - ready_count
        low_confidence_count = int(
            (frame["INCLUDE"].astype(bool) & (frame["CONFIDENCE"] < 0.75)).sum()
        )

        metrics = st.columns(4)
        metrics[0].metric("Candidates", len(frame))
        metrics[1].metric("Included", included_count)
        metrics[2].metric("Ready", ready_count)
        metrics[3].metric("Needs correction", blocked_count)
        if low_confidence_count:
            st.warning(
                f"{low_confidence_count} included row(s) have OCR confidence below 75%. "
                "Check part and item numbers against the scanned page."
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
