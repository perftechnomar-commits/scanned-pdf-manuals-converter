from __future__ import annotations

import io
from pathlib import Path

import pandas as pd
import streamlit as st

from tools import (
    MACHINERY_COLUMNS,
    MACHINERY_TYPES,
    REVIEW_COLUMNS,
    UNIT_OPTIONS,
    build_audit_workbook,
    build_workbook,
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
st.caption(
    "OCR scanned manuals, review the extracted spare-parts rows, and write only "
    "approved records into the original Excel import template."
)


# ---------------------------------------------------------------------------
# Sidebar settings
# ---------------------------------------------------------------------------

with st.sidebar:
    st.header("Source")
    input_type = st.radio(
        "Choose input type",
        ["PDF", "Document URL", "Image", "Image URL"],
        index=0,  # PDF is deliberately the default.
    )

    source_file = None
    document_url = ""
    image_url = ""
    page_spec = ""

    if input_type == "PDF":
        source_file = st.file_uploader("Upload a scanned PDF", type=["pdf"])
        page_spec = st.text_input(
            "Pages to process",
            value="all",
            help="Examples: all, 1-20, 25, 30-35. Process large books in batches.",
        )
    elif input_type == "Document URL":
        document_url = st.text_input("Document URL")
    elif input_type == "Image":
        source_file = st.file_uploader("Upload an image", type=["png", "jpg", "jpeg"])
    else:
        image_url = st.text_input("Image URL")

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
    )
    extraction_model = st.text_input(
        "Structured-extraction model",
        value="mistral-small-latest",
        disabled=structure_mode != "AI JSON extraction (recommended)",
    )
    ocr_pages_per_request = st.number_input(
        "PDF pages per OCR request",
        min_value=1,
        max_value=100,
        value=25,
        step=1,
    )
    extraction_pages_per_batch = st.number_input(
        "OCR pages per structuring batch",
        min_value=1,
        max_value=20,
        value=4,
        step=1,
        disabled=structure_mode != "AI JSON extraction (recommended)",
    )
    default_unit = st.selectbox("Default spare-part unit", ["PCS", "SET", ""], index=0)
    extra_prompt = st.text_area(
        "Optional manual-specific extraction instructions",
        placeholder=(
            "Example: The first column is ITEM NO and the second column is PART NO. "
            "Ignore drawing dimensions and prices."
        ),
        height=100,
        disabled=structure_mode != "AI JSON extraction (recommended)",
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

    if st.button("Reset OCR and review data", use_container_width=True):
        st.session_state.extracted_pages = []
        st.session_state.spare_review = empty_review_dataframe()
        st.session_state.output = None
        st.session_state.editor_version += 1
        st.rerun()


# ---------------------------------------------------------------------------
# Main machinery and optional sub-units
# ---------------------------------------------------------------------------

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
        st.text_input("CODE *", key="main_code", placeholder="M/E")
    with row1[1]:
        st.text_input("NAME *", key="main_name", placeholder="Main Engine")
    with row1[2]:
        st.text_input("MAKER *", key="main_maker", placeholder="Sulzer")
    with row1[3]:
        st.text_input("MODEL *", key="main_model", placeholder="RTA 56")

    row2 = st.columns(3)
    with row2[0]:
        st.text_input("TYPE", key="main_type")
    with row2[1]:
        st.text_input("INSTR.BOOK", key="main_instruction_book")
    with row2[2]:
        st.text_input("SPECIFICATIONS", key="main_specifications")

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

                progress_bar.progress(0.0, text="Converting OCR text into spare-parts rows...")
                extraction_errors: list[str] = []
                if structure_mode == "AI JSON extraction (recommended)":
                    rows, extraction_errors = extract_spare_parts_with_ai(
                        api_key=api_key,
                        model=extraction_model.strip() or "mistral-small-latest",
                        extracted_pages=extracted_pages,
                        pages_per_batch=int(extraction_pages_per_batch),
                        additional_instructions=extra_prompt,
                        progress=show_progress,
                    )
                    if not rows:
                        extraction_errors.append(
                            "AI extraction returned no rows; the local markdown-table parser was used."
                        )
                        rows = extract_spare_parts_from_markdown_tables(extracted_pages)
                else:
                    rows = extract_spare_parts_from_markdown_tables(extracted_pages)

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
                    st.session_state.spare_review = merge_review_dataframes(
                        st.session_state.spare_review,
                        new_review,
                    )
                else:
                    st.session_state.extracted_pages = list(extracted_pages)
                    st.session_state.spare_review = new_review

                st.session_state.editor_version += 1
                st.session_state.output = None
                progress_bar.progress(1.0, text="OCR and extraction complete")
                st.success(
                    f"Processed {len(extracted_pages)} page(s) and created "
                    f"{len(new_review)} candidate spare-part row(s)."
                )
                for message in extraction_errors:
                    st.warning(message)
            except Exception as exc:
                progress_bar.empty()
                st.error(f"Processing failed: {exc}")

    if st.session_state.extracted_pages:
        st.subheader("Raw OCR output")
        page_summary = pd.DataFrame(
            [
                {
                    "Page": page,
                    "Characters": len(markdown),
                    "Preview": markdown[:180].replace("\n", " "),
                }
                for page, markdown in st.session_state.extracted_pages
            ]
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
            output_name = safe_filename(name_source) + "import.xlsx"
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
        )
        st.download_button(
            "Download OCR audit/review workbook",
            data=audit_bytes,
            file_name="OCR_review_and_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
