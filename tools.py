from __future__ import annotations

import io
import json
import os
import re
import tempfile
import time
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import pandas as pd
import requests
from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter


MACHINERY_SHEET = "1.Machineries|Sub|Units"
SPARE_PARTS_SHEET = "2.Spare Parts"

MACHINERY_COLUMNS = [
    "CODE",
    "NAME",
    "MAKER",
    "MODEL",
    "TYPE",
    "INSTR.BOOK",
    "SPECIFICATIONS",
    "MCH_TP(M/S/U)",
]

SPARE_COLUMNS = [
    "MACHINERY",
    "PART NO",
    "DESCRIPTION",
    "CODE",
    "ITEM NO",
    "UNIT",
    "QNT",
]

REVIEW_COLUMNS = [
    "INCLUDE",
    "READY",
    "MACHINERY",
    "PART NO",
    "DESCRIPTION",
    "CODE",
    "ITEM NO",
    "UNIT",
    "QNT",
    "SOURCE PAGE",
    "CONFIDENCE",
    "DETECTED MACHINERY",
    "WARNING",
]

MACHINERY_TYPES = ["Main Machinery", "SubMachinery", "Unit (Book Chapter)"]
UNIT_OPTIONS = ["", "PCS", "SET"]

MAX_MACHINERY_ROWS = 605  # B5:B609 is the template's named machinery range.
MAX_SPARE_ROWS = 1438  # Rows 4:1441 in spare-parts sheet.

ProgressCallback = Callable[[int, int, str], None]


# ---------------------------------------------------------------------------
# General cleanup and validation helpers
# ---------------------------------------------------------------------------


def clean_text(value: Any) -> str:
    """Return a compact, single-line string suitable for review and Excel."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = str(value)
    text = text.replace("\x00", " ")
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"[`*_]+", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_key(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]+", "", clean_text(value).upper())


def excel_safe_text(value: Any) -> str | None:
    """Protect text cells from accidental formula execution while preserving display."""
    text = clean_text(value)
    if not text:
        return None
    if text[0] in ("=", "+", "-", "@"):
        return "'" + text
    return text


def quantity_to_number(value: Any) -> int | float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        try:
            if pd.isna(value):
                return None
        except (TypeError, ValueError):
            pass
        number = float(value)
    else:
        text = clean_text(value).replace(",", ".")
        match = re.search(r"[-+]?\d+(?:\.\d+)?", text)
        if not match:
            return None
        number = float(match.group(0))

    if number.is_integer():
        return int(number)
    return number


def normalize_unit(value: Any, default_unit: str = "PCS") -> str:
    text = clean_text(value).upper()
    if "SET" in text:
        return "SET"
    if any(token in text for token in ("PCS", "PC", "PIECE", "EA", "EACH", "NO.")):
        return "PCS"
    return default_unit if default_unit in UNIT_OPTIONS else ""


def clamp_confidence(value: Any, fallback: float = 0.70) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError):
        number = fallback
    return max(0.0, min(1.0, number))


def empty_review_dataframe() -> pd.DataFrame:
    return pd.DataFrame(columns=REVIEW_COLUMNS)


def empty_additional_machinery_dataframe() -> pd.DataFrame:
    return pd.DataFrame(columns=MACHINERY_COLUMNS)


# ---------------------------------------------------------------------------
# PDF selection and OCR
# ---------------------------------------------------------------------------


def parse_page_spec(spec: str, total_pages: int) -> list[int]:
    """
    Parse an inclusive, one-based page expression such as ``1-20,25,30-35``.

    Returns zero-based page indexes. Blank or ``all`` selects every page.
    """
    if total_pages < 1:
        return []

    text = clean_text(spec).lower()
    if not text or text == "all":
        return list(range(total_pages))

    selected: list[int] = []
    seen: set[int] = set()

    for token in text.split(","):
        token = token.strip()
        if not token:
            continue

        if "-" in token:
            left, right = token.split("-", 1)
            start = int(left) if left.strip() else 1
            end = int(right) if right.strip() else total_pages
            if start > end:
                start, end = end, start
            numbers = range(start, end + 1)
        else:
            numbers = [int(token)]

        for page_number in numbers:
            if not 1 <= page_number <= total_pages:
                raise ValueError(
                    f"Page {page_number} is outside this PDF's range of 1-{total_pages}."
                )
            index = page_number - 1
            if index not in seen:
                selected.append(index)
                seen.add(index)

    if not selected:
        raise ValueError("The page selection did not contain any valid pages.")
    return selected


def pdf_page_count(pdf_bytes: bytes) -> int:
    return len(PdfReader(io.BytesIO(pdf_bytes)).pages)


def chunks(values: Sequence[int], chunk_size: int) -> Iterable[list[int]]:
    chunk_size = max(1, int(chunk_size))
    for start in range(0, len(values), chunk_size):
        yield list(values[start : start + chunk_size])


def _temporary_file(data: bytes, suffix: str) -> str:
    handle = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    try:
        handle.write(data)
        handle.flush()
        return handle.name
    finally:
        handle.close()


def _delete_file(path: str | None) -> None:
    if not path:
        return
    try:
        os.remove(path)
    except FileNotFoundError:
        pass


def ocr_pdf_bytes(
    api_key: str,
    pdf_bytes: bytes,
    page_indexes: Sequence[int] | None = None,
    pages_per_request: int = 25,
    progress: ProgressCallback | None = None,
) -> list[tuple[int, str]]:
    """OCR a PDF in page chunks and preserve original one-based page numbers."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    indexes = list(page_indexes) if page_indexes is not None else list(range(len(reader.pages)))
    if not indexes:
        return []

    from py_mistral_helper.MistralHelper import MistralHelper

    helper = MistralHelper(api_key=api_key)
    page_chunks = list(chunks(indexes, pages_per_request))
    extracted: list[tuple[int, str]] = []

    for chunk_number, page_chunk in enumerate(page_chunks, start=1):
        if progress:
            progress(
                chunk_number - 1,
                len(page_chunks),
                f"OCR request {chunk_number}/{len(page_chunks)}",
            )

        writer = PdfWriter()
        for index in page_chunk:
            writer.add_page(reader.pages[index])

        buffer = io.BytesIO()
        writer.write(buffer)
        temp_path = _temporary_file(buffer.getvalue(), ".pdf")
        try:
            response = helper.extract_text_using_pdf(temp_path)
            response_pages = list(getattr(response, "pages", []) or [])
            for local_index, page in enumerate(response_pages):
                if local_index >= len(page_chunk):
                    break
                original_page = page_chunk[local_index] + 1
                markdown = clean_markdown(getattr(page, "markdown", ""))
                extracted.append((original_page, markdown))
        finally:
            _delete_file(temp_path)

        if progress:
            progress(
                chunk_number,
                len(page_chunks),
                f"OCR request {chunk_number}/{len(page_chunks)} complete",
            )

    extracted.sort(key=lambda item: item[0])
    return extracted


def ocr_document_url(api_key: str, document_url: str) -> list[tuple[int, str]]:
    from py_mistral_helper.MistralHelper import MistralHelper

    helper = MistralHelper(api_key=api_key)
    response = helper.extract_text_using_pdf_document_url(document_url)
    return [
        (index + 1, clean_markdown(getattr(page, "markdown", "")))
        for index, page in enumerate(getattr(response, "pages", []) or [])
    ]


def ocr_image_bytes(api_key: str, image_bytes: bytes, suffix: str) -> list[tuple[int, str]]:
    from py_mistral_helper.MistralHelper import MistralHelper

    helper = MistralHelper(api_key=api_key)
    temp_path = _temporary_file(image_bytes, suffix)
    try:
        response = helper.extract_text_using_image_path(temp_path)
        return [
            (index + 1, clean_markdown(getattr(page, "markdown", "")))
            for index, page in enumerate(getattr(response, "pages", []) or [])
        ]
    finally:
        _delete_file(temp_path)


def ocr_image_url(api_key: str, image_url: str) -> list[tuple[int, str]]:
    from py_mistral_helper.MistralHelper import MistralHelper

    helper = MistralHelper(api_key=api_key)
    response = helper.extract_text_using_image_url(image_url)
    return [
        (index + 1, clean_markdown(getattr(page, "markdown", "")))
        for index, page in enumerate(getattr(response, "pages", []) or [])
    ]


def clean_markdown(value: Any) -> str:
    text = "" if value is None else str(value)
    return text.replace("\x00", "").strip()


# ---------------------------------------------------------------------------
# Structured spare-parts extraction
# ---------------------------------------------------------------------------


EXTRACTION_SYSTEM_PROMPT = """
You are a precise technical-document extraction engine for marine and industrial
spare-parts manuals. Convert OCR markdown into structured spare-parts rows.

Return one JSON object with exactly this top-level key:
{
  "spare_parts": [
    {
      "detected_machinery": "",
      "part_no": "",
      "description": "",
      "code": "",
      "item_no": "",
      "unit": "",
      "quantity": null,
      "source_page": 1,
      "confidence": 0.0
    }
  ]
}

Rules:
1. Extract every genuine spare-part row. Do not summarize and do not invent.
2. Copy part numbers, item numbers, codes, and descriptions exactly as printed,
   apart from removing harmless surrounding whitespace.
3. Keep Part No, Code, and Item No separate. If the document does not provide a
   field, return an empty string.
4. A row is useful when it has a description and at least a part number or item
   number. Do not turn headings, page numbers, drawing labels, or prose into parts.
5. Use the PAGE markers supplied in the user message for source_page.
6. quantity must be a number or null. Do not place text such as AR in quantity.
7. unit should be PCS, SET, or an empty string. Use SET only when explicitly a set.
8. detected_machinery is the closest explicit equipment, assembly, chapter, or
   sub-unit name. Leave it blank if not clear.
9. confidence is a number from 0 to 1 reflecting OCR and row-alignment certainty.
10. Preserve leading zeros and punctuation in identifiers.
""".strip()


def _page_batches(
    extracted_pages: Sequence[tuple[int, str]],
    pages_per_batch: int,
    max_chars: int = 30000,
) -> list[list[tuple[int, str]]]:
    batches: list[list[tuple[int, str]]] = []
    current: list[tuple[int, str]] = []
    current_chars = 0

    for item in extracted_pages:
        page_chars = len(item[1])
        would_overflow = current and (
            len(current) >= max(1, pages_per_batch)
            or current_chars + page_chars > max_chars
        )
        if would_overflow:
            batches.append(current)
            current = []
            current_chars = 0
        current.append(item)
        current_chars += page_chars

    if current:
        batches.append(current)
    return batches


def _mistral_json_request(
    api_key: str,
    model: str,
    messages: list[dict[str, str]],
    timeout_seconds: int = 300,
    max_retries: int = 3,
) -> dict[str, Any]:
    endpoint = os.getenv(
        "MISTRAL_CHAT_COMPLETIONS_URL",
        "https://api.mistral.ai/v1/chat/completions",
    )
    payload = {
        "model": model,
        "messages": messages,
        "temperature": 0,
        "response_format": {"type": "json_object"},
        "max_tokens": 8192,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    last_error: Exception | None = None
    for attempt in range(max_retries):
        try:
            response = requests.post(
                endpoint,
                headers=headers,
                json=payload,
                timeout=timeout_seconds,
            )
            if response.status_code in {429, 500, 502, 503, 504} and attempt < max_retries - 1:
                time.sleep(2 ** (attempt + 1))
                continue
            response.raise_for_status()
            body = response.json()
            content = body["choices"][0]["message"]["content"]
            if isinstance(content, list):
                content = "".join(
                    str(chunk.get("text", ""))
                    for chunk in content
                    if isinstance(chunk, dict)
                )
            text = str(content).strip()
            text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
            text = re.sub(r"\s*```$", "", text)
            parsed = json.loads(text)
            if isinstance(parsed, list):
                return {"spare_parts": parsed}
            if not isinstance(parsed, dict):
                raise ValueError("The structured-extraction response was not a JSON object.")
            return parsed
        except (requests.RequestException, KeyError, IndexError, TypeError, json.JSONDecodeError, ValueError) as exc:
            last_error = exc
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)

    raise RuntimeError(f"Mistral structured extraction failed: {last_error}")


def extract_spare_parts_with_ai(
    api_key: str,
    model: str,
    extracted_pages: Sequence[tuple[int, str]],
    pages_per_batch: int = 4,
    additional_instructions: str = "",
    progress: ProgressCallback | None = None,
) -> tuple[list[dict[str, Any]], list[str]]:
    batches = _page_batches(extracted_pages, pages_per_batch)
    rows: list[dict[str, Any]] = []
    errors: list[str] = []

    for batch_index, batch in enumerate(batches, start=1):
        if progress:
            progress(
                batch_index - 1,
                len(batches),
                f"Structuring OCR batch {batch_index}/{len(batches)}",
            )

        page_text = "\n\n".join(
            f"===== PAGE {page_number} =====\n{markdown}"
            for page_number, markdown in batch
        )
        user_prompt = (
            "Extract all spare-parts rows from the following OCR markdown. "
            "Return only the required JSON object.\n\n"
        )
        if clean_text(additional_instructions):
            user_prompt += (
                "Manual-specific instructions:\n"
                f"{clean_text(additional_instructions)}\n\n"
            )
        user_prompt += page_text

        try:
            result = _mistral_json_request(
                api_key=api_key,
                model=model,
                messages=[
                    {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt},
                ],
            )
            batch_rows = result.get("spare_parts", [])
            if isinstance(batch_rows, list):
                rows.extend(item for item in batch_rows if isinstance(item, dict))
            else:
                errors.append(f"Batch {batch_index}: JSON did not contain a spare_parts list.")
        except Exception as exc:  # Keep remaining batches usable.
            errors.append(f"Batch {batch_index}: {exc}")

        if progress:
            progress(
                batch_index,
                len(batches),
                f"Structuring OCR batch {batch_index}/{len(batches)} complete",
            )

    return rows, errors


# ---------------------------------------------------------------------------
# Local markdown-table fallback
# ---------------------------------------------------------------------------


def _split_markdown_row(line: str) -> list[str]:
    stripped = line.strip().strip("|")
    return [clean_text(cell.replace("\\|", "|")) for cell in re.split(r"(?<!\\)\|", stripped)]


def _is_markdown_separator(line: str) -> bool:
    cells = _split_markdown_row(line)
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells)


def _canonical_header(header: str) -> str | None:
    key = re.sub(r"[^a-z0-9]+", "", header.lower())
    if not key:
        return None
    if any(token in key for token in ("partno", "partnumber", "partnum", "catalogno", "orderingno", "stockno", "pno")):
        return "part_no"
    if any(token in key for token in ("description", "designation", "partname", "denomination")):
        return "description"
    if any(token in key for token in ("itemno", "itemnumber", "positionno", "position", "posno", "refno", "referenceno", "indexno")):
        return "item_no"
    if key in {"code", "partcode", "materialcode", "articlecode"}:
        return "code"
    if any(token in key for token in ("quantity", "qty", "qnt", "nooff", "numberoff")):
        return "quantity"
    if key in {"unit", "uom", "unitofmeasure"}:
        return "unit"
    if any(token in key for token in ("machinery", "equipment", "assembly", "subassembly", "chapter")):
        return "detected_machinery"
    return None


def extract_spare_parts_from_markdown_tables(
    extracted_pages: Sequence[tuple[int, str]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    for page_number, markdown in extracted_pages:
        lines = markdown.splitlines()
        index = 0
        while index + 1 < len(lines):
            header_line = lines[index]
            separator_line = lines[index + 1]
            if "|" not in header_line or not _is_markdown_separator(separator_line):
                index += 1
                continue

            headers = _split_markdown_row(header_line)
            mapped_headers = [_canonical_header(header) for header in headers]
            index += 2

            while index < len(lines) and "|" in lines[index]:
                values = _split_markdown_row(lines[index])
                if len(values) < len(headers):
                    values.extend([""] * (len(headers) - len(values)))
                record: dict[str, Any] = {
                    "source_page": page_number,
                    "confidence": 0.65,
                }
                for column_index, canonical in enumerate(mapped_headers):
                    if canonical and column_index < len(values):
                        record[canonical] = values[column_index]

                if clean_text(record.get("description")) and (
                    clean_text(record.get("part_no"))
                    or clean_text(record.get("item_no"))
                ):
                    rows.append(record)
                index += 1

    return rows


# ---------------------------------------------------------------------------
# Review dataframe construction
# ---------------------------------------------------------------------------


def rows_to_review_dataframe(
    rows: Sequence[dict[str, Any]],
    default_machinery: str,
    default_unit: str = "PCS",
) -> pd.DataFrame:
    normalized: list[dict[str, Any]] = []
    seen: set[tuple[str, ...]] = set()

    for raw in rows:
        part_no = clean_text(raw.get("part_no", raw.get("PART NO", "")))
        description = clean_text(raw.get("description", raw.get("DESCRIPTION", "")))
        code = clean_text(raw.get("code", raw.get("CODE", "")))
        item_no = clean_text(raw.get("item_no", raw.get("ITEM NO", "")))
        detected_machinery = clean_text(
            raw.get("detected_machinery", raw.get("DETECTED MACHINERY", ""))
        )
        source_page = quantity_to_number(raw.get("source_page", raw.get("SOURCE PAGE")))
        source_page_int = int(source_page) if source_page is not None else None
        quantity = quantity_to_number(raw.get("quantity", raw.get("QNT")))
        unit = normalize_unit(raw.get("unit", raw.get("UNIT", "")), default_unit)
        confidence = clamp_confidence(raw.get("confidence", raw.get("CONFIDENCE", 0.70)))

        duplicate_key = (
            str(source_page_int or ""),
            normalize_key(part_no),
            normalize_key(item_no),
            normalize_key(description),
        )
        if duplicate_key in seen:
            continue
        seen.add(duplicate_key)

        useful = bool(description and (part_no or item_no))
        normalized.append(
            {
                "INCLUDE": useful,
                "READY": False,
                "MACHINERY": clean_text(default_machinery),
                "PART NO": part_no,
                "DESCRIPTION": description,
                "CODE": code,
                "ITEM NO": item_no,
                "UNIT": unit,
                "QNT": quantity,
                "SOURCE PAGE": source_page_int,
                "CONFIDENCE": confidence,
                "DETECTED MACHINERY": detected_machinery,
                "WARNING": "",
            }
        )

    if not normalized:
        return empty_review_dataframe()

    frame = pd.DataFrame(normalized, columns=REVIEW_COLUMNS)
    frame["INCLUDE"] = frame["INCLUDE"].astype(bool)
    frame["READY"] = frame["READY"].astype(bool)
    frame["QNT"] = pd.to_numeric(frame["QNT"], errors="coerce")
    frame["SOURCE PAGE"] = pd.to_numeric(frame["SOURCE PAGE"], errors="coerce").astype("Int64")
    frame["CONFIDENCE"] = pd.to_numeric(frame["CONFIDENCE"], errors="coerce").fillna(0.70)
    return frame


def merge_review_dataframes(existing: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    if existing is None or existing.empty:
        return new.copy()
    if new is None or new.empty:
        return existing.copy()

    combined = pd.concat([existing, new], ignore_index=True)
    keys = combined.apply(
        lambda row: "|".join(
            [
                str(row.get("SOURCE PAGE", "")),
                normalize_key(row.get("PART NO", "")),
                normalize_key(row.get("ITEM NO", "")),
                normalize_key(row.get("DESCRIPTION", "")),
            ]
        ),
        axis=1,
    )
    return combined.loc[~keys.duplicated()].reset_index(drop=True)


def machinery_rows_from_main_and_additional(
    main_row: dict[str, Any],
    additional: pd.DataFrame | None,
) -> pd.DataFrame:
    records = [{column: clean_text(main_row.get(column, "")) for column in MACHINERY_COLUMNS}]
    if additional is not None and not additional.empty:
        for _, row in additional.iterrows():
            record = {column: clean_text(row.get(column, "")) for column in MACHINERY_COLUMNS}
            if any(record.values()):
                records.append(record)
    return pd.DataFrame(records, columns=MACHINERY_COLUMNS)


def validate_machinery_dataframe(frame: pd.DataFrame) -> list[str]:
    errors: list[str] = []
    if frame is None or frame.empty:
        return ["At least one main machinery row is required."]
    if len(frame) > MAX_MACHINERY_ROWS:
        errors.append(
            f"The template supports at most {MAX_MACHINERY_ROWS} machinery rows."
        )

    required = ["CODE", "NAME", "MAKER", "MODEL", "MCH_TP(M/S/U)"]
    for index, row in frame.reset_index(drop=True).iterrows():
        excel_row = index + 5
        for column in required:
            if not clean_text(row.get(column, "")):
                errors.append(f"Machinery sheet row {excel_row}: {column} is required.")
        machinery_type = clean_text(row.get("MCH_TP(M/S/U)", ""))
        if machinery_type and machinery_type not in MACHINERY_TYPES:
            errors.append(
                f"Machinery sheet row {excel_row}: invalid MCH_TP value '{machinery_type}'."
            )

    main_count = sum(
        clean_text(value) == "Main Machinery"
        for value in frame["MCH_TP(M/S/U)"].tolist()
    )
    if main_count != 1:
        errors.append("The workbook must contain exactly one Main Machinery row.")

    names = [normalize_key(value) for value in frame["NAME"].tolist() if clean_text(value)]
    duplicate_names = {name for name in names if names.count(name) > 1}
    if duplicate_names:
        errors.append("Machinery NAME values must be unique.")

    codes = [normalize_key(value) for value in frame["CODE"].tolist() if clean_text(value)]
    duplicate_codes = {code for code in codes if codes.count(code) > 1}
    if duplicate_codes:
        errors.append("Machinery CODE values must be unique.")
    return errors


def recalculate_review_status(
    frame: pd.DataFrame,
    valid_machinery_names: Sequence[str],
    allow_duplicates: bool = False,
) -> pd.DataFrame:
    if frame is None or frame.empty:
        return empty_review_dataframe()

    result = frame.copy()
    for column in REVIEW_COLUMNS:
        if column not in result.columns:
            result[column] = False if column in {"INCLUDE", "READY"} else ""

    valid_names = {normalize_key(name) for name in valid_machinery_names if clean_text(name)}
    included_indexes = [index for index, value in result["INCLUDE"].items() if bool(value)]

    duplicate_counter: dict[tuple[str, str, str, str], int] = {}
    for index in included_indexes:
        row = result.loc[index]
        key = (
            normalize_key(row.get("MACHINERY", "")),
            normalize_key(row.get("PART NO", "")),
            normalize_key(row.get("ITEM NO", "")),
            normalize_key(row.get("DESCRIPTION", "")),
        )
        duplicate_counter[key] = duplicate_counter.get(key, 0) + 1

    warnings: list[str] = []
    ready_values: list[bool] = []

    for index, row in result.iterrows():
        if not bool(row.get("INCLUDE", False)):
            warnings.append("Excluded from export")
            ready_values.append(False)
            continue

        row_messages: list[str] = []
        blocking = False
        machinery = clean_text(row.get("MACHINERY", ""))
        part_no = clean_text(row.get("PART NO", ""))
        description = clean_text(row.get("DESCRIPTION", ""))
        item_no = clean_text(row.get("ITEM NO", ""))
        unit = clean_text(row.get("UNIT", "")).upper()

        if not machinery:
            row_messages.append("Missing machinery")
            blocking = True
        elif normalize_key(machinery) not in valid_names:
            row_messages.append("Machinery is not present on sheet 1")
            blocking = True
        if not description:
            row_messages.append("Missing description")
            blocking = True
        if not part_no and not item_no:
            row_messages.append("Part No or Item No is required")
            blocking = True
        if unit not in UNIT_OPTIONS:
            row_messages.append("Unit must be blank, PCS, or SET")
            blocking = True

        raw_quantity = row.get("QNT")
        if clean_text(raw_quantity) and quantity_to_number(raw_quantity) is None:
            row_messages.append("Quantity is not numeric")
            blocking = True

        duplicate_key = (
            normalize_key(machinery),
            normalize_key(part_no),
            normalize_key(item_no),
            normalize_key(description),
        )
        if duplicate_counter.get(duplicate_key, 0) > 1:
            row_messages.append("Possible duplicate")
            if not allow_duplicates:
                blocking = True

        confidence = clamp_confidence(row.get("CONFIDENCE", 0.70))
        if confidence < 0.75:
            row_messages.append("Low OCR confidence - review identifiers")

        warnings.append("; ".join(row_messages))
        ready_values.append(not blocking)

    result["WARNING"] = warnings
    result["READY"] = ready_values
    result["QNT"] = pd.to_numeric(result["QNT"], errors="coerce")
    result["SOURCE PAGE"] = pd.to_numeric(result["SOURCE PAGE"], errors="coerce").astype("Int64")
    result["CONFIDENCE"] = pd.to_numeric(result["CONFIDENCE"], errors="coerce").fillna(0.70)
    return result[REVIEW_COLUMNS]


# ---------------------------------------------------------------------------
# Template and audit workbook generation
# ---------------------------------------------------------------------------


def _clear_values(worksheet: Any, min_row: int, max_row: int, min_col: int, max_col: int) -> None:
    for row in worksheet.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
    ):
        for cell in row:
            cell.value = None


def build_workbook(
    template_bytes: bytes,
    machinery_frame: pd.DataFrame,
    review_frame: pd.DataFrame,
    clear_existing: bool = True,
) -> bytes:
    if len(machinery_frame) > MAX_MACHINERY_ROWS:
        raise ValueError(f"Too many machinery rows; maximum is {MAX_MACHINERY_ROWS}.")

    selected = review_frame[
        review_frame["INCLUDE"].astype(bool) & review_frame["READY"].astype(bool)
    ].copy()
    if len(selected) > MAX_SPARE_ROWS:
        raise ValueError(f"Too many spare-parts rows; maximum is {MAX_SPARE_ROWS}.")

    workbook = load_workbook(io.BytesIO(template_bytes), data_only=False, keep_links=True)
    missing_sheets = [
        name for name in (MACHINERY_SHEET, SPARE_PARTS_SHEET) if name not in workbook.sheetnames
    ]
    if missing_sheets:
        raise ValueError(
            "The selected workbook is not the expected template. Missing sheets: "
            + ", ".join(missing_sheets)
        )

    machinery_sheet = workbook[MACHINERY_SHEET]
    spare_sheet = workbook[SPARE_PARTS_SHEET]

    if clear_existing:
        _clear_values(machinery_sheet, 5, 609, 1, 8)
        _clear_values(spare_sheet, 4, 1441, 1, 7)

    for offset, (_, row) in enumerate(machinery_frame.iterrows(), start=5):
        for column_index, column_name in enumerate(MACHINERY_COLUMNS, start=1):
            cell = machinery_sheet.cell(row=offset, column=column_index)
            cell.value = excel_safe_text(row.get(column_name, ""))
            cell.number_format = "@"

    for offset, (_, row) in enumerate(selected.iterrows(), start=4):
        values = [
            row.get("MACHINERY", ""),
            row.get("PART NO", ""),
            row.get("DESCRIPTION", ""),
            row.get("CODE", ""),
            row.get("ITEM NO", ""),
            normalize_unit(row.get("UNIT", ""), ""),
            quantity_to_number(row.get("QNT")),
        ]
        for column_index, value in enumerate(values, start=1):
            cell = spare_sheet.cell(row=offset, column=column_index)
            if column_index == 7:
                cell.value = value
                cell.number_format = "0.###"
            else:
                cell.value = excel_safe_text(value)
                cell.number_format = "@"

    workbook.active = workbook.sheetnames.index(SPARE_PARTS_SHEET)
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def build_audit_workbook(
    extracted_pages: Sequence[tuple[int, str]],
    machinery_frame: pd.DataFrame,
    review_frame: pd.DataFrame,
) -> bytes:
    output = io.BytesIO()
    pages_frame = pd.DataFrame(extracted_pages, columns=["SOURCE PAGE", "OCR MARKDOWN"])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pages_frame.to_excel(writer, index=False, sheet_name="OCR Pages")
        machinery_frame.to_excel(writer, index=False, sheet_name="Machinery Review")
        review_frame.to_excel(writer, index=False, sheet_name="Spare Parts Review")

        workbook = writer.book
        header_format = workbook.add_format(
            {"bold": True, "bg_color": "#D9EAF7", "border": 1}
        )
        wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})

        for sheet_name, frame in (
            ("OCR Pages", pages_frame),
            ("Machinery Review", machinery_frame),
            ("Spare Parts Review", review_frame),
        ):
            sheet = writer.sheets[sheet_name]
            for column_index, column_name in enumerate(frame.columns):
                sheet.write(0, column_index, column_name, header_format)
            sheet.freeze_panes(1, 0)
            sheet.autofilter(0, 0, max(1, len(frame)), max(0, len(frame.columns) - 1))

        writer.sheets["OCR Pages"].set_column(0, 0, 12)
        writer.sheets["OCR Pages"].set_column(1, 1, 100, wrap_format)
        writer.sheets["Machinery Review"].set_column(0, len(MACHINERY_COLUMNS) - 1, 22)
        writer.sheets["Spare Parts Review"].set_column(0, len(REVIEW_COLUMNS) - 1, 20)
        writer.sheets["Spare Parts Review"].set_column(
            REVIEW_COLUMNS.index("DESCRIPTION"),
            REVIEW_COLUMNS.index("DESCRIPTION"),
            50,
            wrap_format,
        )
        writer.sheets["Spare Parts Review"].set_column(
            REVIEW_COLUMNS.index("WARNING"),
            REVIEW_COLUMNS.index("WARNING"),
            45,
            wrap_format,
        )

    return output.getvalue()


def safe_filename(value: str, fallback: str = "spare_parts") -> str:
    stem = Path(clean_text(value)).stem
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._-")
    return stem or fallback
