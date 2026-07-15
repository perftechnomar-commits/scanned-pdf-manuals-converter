from __future__ import annotations

import io
import json
import os
import re
import tempfile
import time
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence
from difflib import SequenceMatcher
from datetime import datetime, timezone

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

BENEFIT_SPARE_COLUMNS = [
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
    "SECTION START PAGE",
    "TABLE TITLE",
    "CONFIDENCE",
    "DETECTED MACHINERY",
    "ASSIGNMENT SOURCE",
    "WARNING",
]

SUBMACHINERY_REVIEW_COLUMNS = [
    "INCLUDE",
    "CODE",
    "NAME",
    "MAKER",
    "MODEL",
    "TYPE",
    "INSTR.BOOK",
    "SPECIFICATIONS",
    "MCH_TP(M/S/U)",
    "FIRST PAGE",
    "LAST PAGE",
    "PARTS FOUND",
    "CONFIDENCE",
    "VARIANTS",
    "DETECTION KEYS",
    "ORIGIN",
]

MACHINERY_TYPES = ["Main Machinery", "SubMachinery"]
UNIT_OPTIONS = ["", "PCS", "SET"]

MAX_MACHINERY_ROWS = 605  # B5:B609 is the template's named machinery range.
MAX_SPARE_ROWS = 1438  # Rows 4:1441 in the spare-parts import sheet.

ProgressCallback = Callable[[int, int, str], None]

PAGE_CLASSIFICATION_COLUMNS = [
    "SOURCE PAGE",
    "CLASSIFICATION",
    "PROCESS",
    "SCORE",
    "REASON",
    "CHARACTERS",
]

PAGE_FILTER_MODES = [
    "Conservative (recommended)",
    "Strict",
    "Off",
]


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


def empty_submachinery_review_dataframe() -> pd.DataFrame:
    return pd.DataFrame(columns=SUBMACHINERY_REVIEW_COLUMNS)


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
            response = None
            retry_delays = (0, 2, 4, 8, 12)
            last_error: Exception | None = None

            for attempt, delay_seconds in enumerate(retry_delays, start=1):
                if delay_seconds:
                    if progress:
                        progress(
                            chunk_number - 1,
                            len(page_chunks),
                            (
                                f"OCR request {chunk_number}/{len(page_chunks)}: "
                                f"waiting {delay_seconds}s before retry {attempt}/{len(retry_delays)}"
                            ),
                        )
                    time.sleep(delay_seconds)

                try:
                    # Recreate the helper for every retry. This avoids reusing stale
                    # uploaded-file state inside py-mistral-helper after Mistral's file
                    # service temporarily returns: 404 "No file matches the given query".
                    attempt_helper = MistralHelper(api_key=api_key)
                    response = attempt_helper.extract_text_using_pdf(temp_path)
                    break
                except Exception as exc:
                    last_error = exc
                    message = str(exc).lower()
                    transient_file_error = (
                        "no file matches the given query" in message
                        or "status 404" in message
                        or "status_code=404" in message
                    )

                    if not transient_file_error or attempt == len(retry_delays):
                        raise RuntimeError(
                            f"OCR request {chunk_number}/{len(page_chunks)} failed "
                            f"after {attempt} attempt(s): {exc}"
                        ) from exc

            if response is None:
                raise RuntimeError(
                    f"OCR request {chunk_number}/{len(page_chunks)} returned no response: "
                    f"{last_error}"
                )

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
# Local page classification / pre-filtering
# ---------------------------------------------------------------------------


_POSITIVE_PAGE_PHRASES: tuple[tuple[str, int], ...] = (
    ("spare parts", 5),
    ("list of parts", 5),
    ("parts list", 5),
    ("part no", 4),
    ("part number", 4),
    ("item no", 3),
    ("item number", 3),
    ("position no", 3),
    ("description", 2),
    ("designation", 2),
    ("denomination", 2),
    ("quantity", 2),
    (" qty", 2),
    (" qnt", 2),
)

_NEGATIVE_PAGE_PHRASES: tuple[tuple[str, int], ...] = (
    ("table of contents", 7),
    ("revision history", 6),
    ("list of revisions", 6),
    ("document revisions", 6),
    ("foreword", 4),
    ("preface", 4),
    ("general description", 3),
    ("operating instructions", 3),
    ("safety instructions", 3),
    ("maintenance instructions", 2),
)

_PART_IDENTIFIER_PATTERN = re.compile(
    r"(?<![A-Z0-9])[A-Z0-9][A-Z0-9./_-]{3,}(?![A-Z0-9])",
    flags=re.IGNORECASE,
)


def _markdown_table_signal(markdown: str) -> tuple[int, int]:
    """Return (table rows, relevant-header hits) for OCR markdown."""
    lines = markdown.splitlines()
    table_rows = sum(1 for line in lines if line.count("|") >= 2)
    header_hits = 0
    for line in lines:
        if "|" not in line:
            continue
        lowered = line.lower()
        if any(
            phrase in lowered
            for phrase in (
                "part no",
                "part number",
                "item no",
                "item number",
                "description",
                "designation",
                "quantity",
                "qty",
                "qnt",
            )
        ):
            header_hits += 1
    return table_rows, header_hits


def classify_ocr_pages(
    extracted_pages: Sequence[tuple[int, str]],
    mode: str = "Conservative (recommended)",
) -> tuple[list[tuple[int, str]], pd.DataFrame]:
    """
    Classify OCR pages locally before paid structured extraction.

    Conservative mode skips only pages that are very clearly front matter or prose.
    Strict mode processes only strong spare-parts candidates. Off processes everything.
    No extra API calls are made.
    """
    if mode not in PAGE_FILTER_MODES:
        mode = "Conservative (recommended)"

    selected: list[tuple[int, str]] = []
    records: list[dict[str, Any]] = []

    for page_number, markdown in extracted_pages:
        text = clean_markdown(markdown)
        lowered = text.lower()
        characters = len(text)
        table_rows, header_hits = _markdown_table_signal(text)
        identifier_hits = len(_PART_IDENTIFIER_PATTERN.findall(text))

        positive_score = sum(weight for phrase, weight in _POSITIVE_PAGE_PHRASES if phrase in lowered)
        negative_score = sum(weight for phrase, weight in _NEGATIVE_PAGE_PHRASES if phrase in lowered)
        score = positive_score + min(table_rows, 8) + (header_hits * 3) + min(identifier_hits // 4, 5) - negative_score

        strong_candidate = (
            positive_score >= 5
            or header_hits >= 1
            or (table_rows >= 4 and identifier_hits >= 3)
        )
        obvious_non_parts = (
            characters < 40
            or (negative_score >= 5 and not strong_candidate)
            or (table_rows == 0 and positive_score == 0 and identifier_hits < 2 and characters < 1800)
        )

        if mode == "Off":
            classification = "All pages"
            process_page = True
            reason = "Filtering disabled"
        elif strong_candidate:
            classification = "Spare-parts candidate"
            process_page = True
            reason = (
                f"parts signals={positive_score}; table rows={table_rows}; "
                f"header hits={header_hits}; identifiers={identifier_hits}"
            )
        elif obvious_non_parts:
            classification = "Skipped obvious non-parts page"
            process_page = False
            reason = (
                f"negative signals={negative_score}; table rows={table_rows}; "
                f"parts signals={positive_score}; identifiers={identifier_hits}"
            )
        else:
            classification = "Ambiguous"
            process_page = mode == "Conservative (recommended)"
            reason = (
                f"score={score}; table rows={table_rows}; "
                f"parts signals={positive_score}; identifiers={identifier_hits}"
            )

        if process_page:
            selected.append((page_number, text))

        records.append(
            {
                "SOURCE PAGE": int(page_number),
                "CLASSIFICATION": classification,
                "PROCESS": bool(process_page),
                "SCORE": int(score),
                "REASON": reason,
                "CHARACTERS": int(characters),
            }
        )

    return selected, pd.DataFrame(records, columns=PAGE_CLASSIFICATION_COLUMNS)


# ---------------------------------------------------------------------------
# Structured spare-parts extraction
# ---------------------------------------------------------------------------


EXTRACTION_SYSTEM_PROMPT = """
You are a precise technical-document extraction engine for marine and industrial
spare-parts manuals. Convert OCR markdown into structured spare-parts rows and
identify the sub-machinery or assembly heading that governs each table.

Return one JSON object with exactly this top-level key:
{
  "spare_parts": [
    {
      "detected_machinery": "",
      "table_title": "",
      "section_start_page": 1,
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
2. Copy part numbers, item numbers, codes, descriptions, and headings exactly as
   printed, apart from harmless surrounding whitespace.
3. Keep Part No, Code, and Item No separate. If the document does not provide a
   field, return an empty string.
4. A row is useful when it has a description and at least a part number or item
   number. Do not turn headings, page numbers, drawing labels, or prose into parts.
5. Use the PAGE markers supplied in the user message for source_page.
6. quantity must be a number or null. Do not place text such as AR in quantity.
7. unit should be PCS, SET, or an empty string. Use SET only when explicitly a set.
8. detected_machinery must be the closest explicit equipment, assembly, component,
   or sub-machinery title printed above or immediately before the parts table. Do
   not use generic text such as SPARE PARTS LIST, PARTS CATALOGUE, DESCRIPTION, or
   the main manual title as detected_machinery.
9. table_title is the complete table/section heading as printed. It may be the same
   as detected_machinery.
10. Repeat detected_machinery and table_title for every row belonging to the same
    table, including continuation pages where the heading is not repeated. Use the
    nearest preceding heading within the supplied pages when the table continues.
11. section_start_page is the first supplied page on which that table or section
    begins. Keep it the same for all continuation rows in that section.
12. confidence is a number from 0 to 1 reflecting OCR, heading detection, and row
    alignment certainty.
13. Preserve leading zeros and punctuation in identifiers.
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


def _parse_json_object(content: Any) -> dict[str, Any]:
    """Parse a JSON object while tolerating fences or harmless surrounding text."""
    if isinstance(content, list):
        content = "".join(
            str(chunk.get("text", ""))
            for chunk in content
            if isinstance(chunk, dict)
        )
    text = str(content or "").strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)

    candidates: list[str] = [text]
    first_brace = text.find("{")
    last_brace = text.rfind("}")
    if first_brace >= 0 and last_brace > first_brace:
        candidates.append(text[first_brace : last_brace + 1])

    decoder = json.JSONDecoder()
    for candidate in candidates:
        if not candidate:
            continue
        for normalized in (
            candidate,
            re.sub(r",\s*([}\]])", r"\1", candidate),
        ):
            try:
                parsed = json.loads(normalized, strict=False)
            except json.JSONDecodeError:
                try:
                    parsed, _ = decoder.raw_decode(normalized)
                except json.JSONDecodeError:
                    continue
            if isinstance(parsed, list):
                return {"spare_parts": parsed}
            if isinstance(parsed, dict):
                return parsed

    raise json.JSONDecodeError("Could not parse a complete JSON object", text, 0)


def _mistral_json_request(
    api_key: str,
    model: str,
    messages: list[dict[str, str]],
    timeout_seconds: int = 300,
    max_retries: int = 2,
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
            return _parse_json_object(body["choices"][0]["message"]["content"])
        except (
            requests.RequestException,
            KeyError,
            IndexError,
            TypeError,
            json.JSONDecodeError,
            ValueError,
        ) as exc:
            last_error = exc
            if attempt < max_retries - 1:
                time.sleep(1 + attempt)

    raise RuntimeError(f"Mistral structured extraction failed: {last_error}")


def _build_extraction_prompt(
    batch: Sequence[tuple[int, str]],
    additional_instructions: str,
) -> str:
    page_text = "\n\n".join(
        f"===== PAGE {page_number} =====\n{markdown}"
        for page_number, markdown in batch
    )
    prompt = (
        "Extract all genuine spare-parts rows from the following OCR markdown. "
        "Return only the required JSON object. If the pages contain no spare-parts "
        "rows, return {\"spare_parts\": []}.\n\n"
    )
    if clean_text(additional_instructions):
        prompt += (
            "Manual-specific instructions:\n"
            f"{clean_text(additional_instructions)}\n\n"
        )
    return prompt + page_text


def _normalize_batch_source_pages(
    batch_rows: Sequence[dict[str, Any]],
    batch: Sequence[tuple[int, str]],
) -> list[dict[str, Any]]:
    valid_pages = {int(page) for page, _ in batch}
    only_page = next(iter(valid_pages)) if len(valid_pages) == 1 else None
    normalized: list[dict[str, Any]] = []
    for item in batch_rows:
        if not isinstance(item, dict):
            continue
        row = dict(item)
        parsed_page = quantity_to_number(row.get("source_page"))
        page_number = int(parsed_page) if parsed_page is not None else None
        if page_number not in valid_pages and only_page is not None:
            page_number = only_page
        if page_number is not None:
            row["source_page"] = page_number

        parsed_start = quantity_to_number(row.get("section_start_page"))
        section_start = int(parsed_start) if parsed_start is not None else None
        if section_start not in valid_pages:
            section_start = page_number
        if section_start is not None and page_number is not None:
            section_start = min(section_start, page_number)
        row["section_start_page"] = section_start

        detected = clean_text(row.get("detected_machinery"))
        table_title = clean_text(row.get("table_title"))
        if not detected and table_title and not _is_generic_machinery_name(table_title):
            row["detected_machinery"] = table_title
        normalized.append(row)
    return normalized


def _propagate_detected_machinery_context(
    rows: Sequence[dict[str, Any]],
    max_page_gap: int = 2,
) -> list[dict[str, Any]]:
    """Carry a clear section heading onto nearby continuation rows conservatively."""
    result = [dict(row) for row in rows if isinstance(row, dict)]
    indexed = sorted(
        enumerate(result),
        key=lambda item: (
            quantity_to_number(item[1].get("source_page")) or 10**9,
            item[0],
        ),
    )
    last_detected = ""
    last_title = ""
    last_section_page: int | None = None
    last_source_page: int | None = None

    for _, row in indexed:
        source_value = quantity_to_number(row.get("source_page"))
        source_page = int(source_value) if source_value is not None else None
        detected = clean_text(row.get("detected_machinery"))
        title = clean_text(row.get("table_title"))
        start_value = quantity_to_number(row.get("section_start_page"))
        section_page = int(start_value) if start_value is not None else None

        if detected and not _is_generic_machinery_name(detected):
            last_detected = detected
            last_title = title or detected
            last_section_page = section_page or source_page
            last_source_page = source_page
            continue

        nearby = (
            source_page is not None
            and last_source_page is not None
            and 0 <= source_page - last_source_page <= max_page_gap
        )
        if nearby and last_detected:
            row["detected_machinery"] = last_detected
            if not title:
                row["table_title"] = last_title
            if section_page is None:
                row["section_start_page"] = last_section_page
            row["machinery_inherited"] = True
            last_source_page = source_page

    return result


def extract_spare_parts_with_ai(
    api_key: str,
    model: str,
    extracted_pages: Sequence[tuple[int, str]],
    pages_per_batch: int = 3,
    max_chars_per_batch: int = 12000,
    additional_instructions: str = "",
    progress: ProgressCallback | None = None,
) -> tuple[list[dict[str, Any]], list[str]]:
    """
    Extract rows with automatic divide-and-retry recovery.

    If a multi-page response is malformed, the batch is split in half recursively.
    If a single page still fails, the local markdown-table parser is used for that
    page so one bad response does not discard the remainder of a large manual.
    """
    batches = _page_batches(
        extracted_pages,
        pages_per_batch=max(1, int(pages_per_batch)),
        max_chars=max(2000, int(max_chars_per_batch)),
    )
    rows: list[dict[str, Any]] = []
    messages: list[str] = []

    def process_batch(
        batch: list[tuple[int, str]],
        label: str,
        depth: int = 0,
    ) -> None:
        try:
            result = _mistral_json_request(
                api_key=api_key,
                model=model,
                messages=[
                    {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                    {
                        "role": "user",
                        "content": _build_extraction_prompt(batch, additional_instructions),
                    },
                ],
            )
            batch_rows = result.get("spare_parts", [])
            if not isinstance(batch_rows, list):
                raise ValueError("JSON did not contain a spare_parts list")
            rows.extend(_normalize_batch_source_pages(batch_rows, batch))
            return
        except Exception as exc:
            if len(batch) > 1:
                midpoint = max(1, len(batch) // 2)
                left = batch[:midpoint]
                right = batch[midpoint:]
                page_range = f"{batch[0][0]}-{batch[-1][0]}"
                messages.append(
                    f"Recovered {label} (pages {page_range}) by automatically splitting "
                    "the malformed/oversized response into smaller requests."
                )
                process_batch(left, f"{label}.1", depth + 1)
                process_batch(right, f"{label}.2", depth + 1)
                return

            page_number = batch[0][0]
            fallback_rows = extract_spare_parts_from_markdown_tables(batch)
            rows.extend(fallback_rows)
            if fallback_rows:
                messages.append(
                    f"Page {page_number}: AI JSON remained invalid; the local table "
                    f"parser recovered {len(fallback_rows)} row(s)."
                )
            else:
                messages.append(
                    f"Page {page_number}: structured extraction failed and no local "
                    f"table rows were recoverable. Details: {exc}"
                )

    for batch_index, batch in enumerate(batches, start=1):
        if progress:
            progress(
                batch_index - 1,
                len(batches),
                f"Structuring candidate batch {batch_index}/{len(batches)}",
            )
        process_batch(list(batch), f"batch {batch_index}")
        if progress:
            progress(
                batch_index,
                len(batches),
                f"Structuring candidate batch {batch_index}/{len(batches)} complete",
            )

    # Carry clear headings onto nearby continuation pages, then avoid duplicate
    # recovery messages when recursive splitting happened repeatedly.
    rows = _propagate_detected_machinery_context(rows)
    deduplicated_messages = list(dict.fromkeys(messages))
    return rows, deduplicated_messages


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


def _nearest_table_title(lines: Sequence[str], header_index: int) -> str:
    """Return the nearest plausible heading before a Markdown table."""
    for candidate_index in range(header_index - 1, max(-1, header_index - 10), -1):
        candidate = clean_text(lines[candidate_index].lstrip("# "))
        if not candidate or "|" in candidate:
            continue
        if len(candidate) < 3 or len(candidate) > 140:
            continue
        if re.fullmatch(r"(?:page\s*)?\d+(?:\s*/\s*\d+)?", candidate, flags=re.I):
            continue
        if _is_generic_machinery_name(candidate):
            continue
        return candidate
    return ""


def extract_spare_parts_from_markdown_tables(
    extracted_pages: Sequence[tuple[int, str]],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    for page_number, markdown in extracted_pages:
        lines = markdown.splitlines()
        index = 0
        while index + 1 < len(lines):
            header_index = index
            header_line = lines[index]
            separator_line = lines[index + 1]
            if "|" not in header_line or not _is_markdown_separator(separator_line):
                index += 1
                continue

            table_title = _nearest_table_title(lines, header_index)
            headers = _split_markdown_row(header_line)
            mapped_headers = [_canonical_header(header) for header in headers]
            index += 2

            while index < len(lines) and "|" in lines[index]:
                values = _split_markdown_row(lines[index])
                if len(values) < len(headers):
                    values.extend([""] * (len(headers) - len(values)))
                record: dict[str, Any] = {
                    "source_page": page_number,
                    "section_start_page": page_number,
                    "table_title": table_title,
                    "detected_machinery": table_title,
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

    return _propagate_detected_machinery_context(rows)



# ---------------------------------------------------------------------------
# Automatic sub-machinery detection and assignment
# ---------------------------------------------------------------------------


_GENERIC_MACHINERY_NAMES = {
    "SPARE PARTS",
    "SPARE PARTS LIST",
    "PARTS LIST",
    "LIST OF PARTS",
    "PARTS CATALOGUE",
    "PARTS CATALOG",
    "CATALOGUE",
    "CATALOG",
    "DESCRIPTION",
    "DESIGNATION",
    "ITEM NO",
    "ITEM NUMBER",
    "PART NO",
    "PART NUMBER",
    "QUANTITY",
    "DRAWING",
    "DRAWING NO",
    "TABLE",
    "CONTINUED",
}


def _clean_machinery_name(value: Any) -> str:
    text = clean_text(value).strip(" -:;|/")
    text = re.sub(
        r"^(?:spare\s+parts(?:\s+list)?|parts\s+list|list\s+of\s+parts)\s*(?:for|of)?\s*[:\-]*\s*",
        "",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(r"\s+(?:continued|cont\.?)(?:\s*\(.*?\))?$", "", text, flags=re.I)
    text = re.sub(r"\s+page\s+\d+(?:\s+of\s+\d+)?$", "", text, flags=re.I)
    return clean_text(text).strip(" -:;|/")


def _is_generic_machinery_name(value: Any) -> bool:
    text = _clean_machinery_name(value)
    if not text or len(text) < 3:
        return True
    normalized = re.sub(r"[^A-Z0-9 ]+", " ", text.upper())
    normalized = re.sub(r"\s+", " ", normalized).strip()
    if normalized in _GENERIC_MACHINERY_NAMES:
        return True
    if re.fullmatch(r"(?:FIG(?:URE)?|DWG|DRAWING|PAGE|TABLE)\s*[A-Z0-9./_-]*", normalized):
        return True
    if not re.search(r"[A-Z]", normalized):
        return True
    return False


def _machinery_name_similarity(left: str, right: str) -> float:
    left_clean = _clean_machinery_name(left).upper()
    right_clean = _clean_machinery_name(right).upper()
    if not left_clean or not right_clean:
        return 0.0
    left_key = normalize_key(left_clean)
    right_key = normalize_key(right_clean)
    if left_key == right_key:
        return 1.0

    sequence = SequenceMatcher(None, left_key, right_key).ratio()
    left_tokens = {token for token in re.findall(r"[A-Z0-9]+", left_clean) if len(token) > 1}
    right_tokens = {token for token in re.findall(r"[A-Z0-9]+", right_clean) if len(token) > 1}
    token_score = (
        len(left_tokens & right_tokens) / max(1, len(left_tokens | right_tokens))
    )
    # Avoid merging a component with a more specific component merely because one
    # name contains the other (for example FUEL PUMP vs FUEL PUMP COVER). Only
    # tolerate harmless generic suffix differences.
    generic_suffixes = {"ASSEMBLY", "ASSY", "UNIT", "COMPLETE"}
    containment = 0.0
    if min(len(left_key), len(right_key)) >= 8 and (
        left_key in right_key or right_key in left_key
    ):
        extra_tokens = (left_tokens | right_tokens) - (left_tokens & right_tokens)
        if extra_tokens and extra_tokens <= generic_suffixes:
            containment = 0.94
    return max(sequence, token_score, containment)


def _split_detection_keys(value: Any) -> set[str]:
    return {
        key.strip()
        for key in clean_text(value).split("|")
        if key.strip()
    }


def _generated_submachinery_code(position: int, existing_codes: set[str]) -> str:
    counter = max(1, int(position))
    while True:
        candidate = f"SUB-{counter:03d}"
        if normalize_key(candidate) not in existing_codes:
            existing_codes.add(normalize_key(candidate))
            return candidate
        counter += 1


def build_submachinery_candidates(
    review_frame: pd.DataFrame,
    main_row: dict[str, Any],
) -> pd.DataFrame:
    """Create editable sub-machinery proposals from detected table headings."""
    if review_frame is None or review_frame.empty:
        return empty_submachinery_review_dataframe()

    main_name = clean_text(main_row.get("NAME", ""))
    observations: list[dict[str, Any]] = []
    for _, row in review_frame.iterrows():
        detected = _clean_machinery_name(
            row.get("DETECTED MACHINERY", row.get("TABLE TITLE", ""))
        )
        if _is_generic_machinery_name(detected):
            continue
        if main_name and _machinery_name_similarity(detected, main_name) >= 0.96:
            continue

        source = quantity_to_number(row.get("SOURCE PAGE"))
        section = quantity_to_number(row.get("SECTION START PAGE"))
        confidence = clamp_confidence(row.get("CONFIDENCE", 0.70))
        key = normalize_key(detected)
        if not key:
            continue
        observations.append(
            {
                "name": detected,
                "key": key,
                "source_page": int(source) if source is not None else None,
                "section_page": int(section) if section is not None else None,
                "confidence": confidence,
            }
        )

    if not observations:
        return empty_submachinery_review_dataframe()

    groups: list[dict[str, Any]] = []
    for observation in observations:
        best_group: dict[str, Any] | None = None
        best_score = 0.0
        for group in groups:
            score = _machinery_name_similarity(observation["name"], group["representative"])
            if score > best_score:
                best_score = score
                best_group = group
        if best_group is not None and best_score >= 0.90:
            best_group["observations"].append(observation)
            # Prefer the longer, more descriptive variant as representative.
            if len(observation["name"]) > len(best_group["representative"]):
                best_group["representative"] = observation["name"]
        else:
            groups.append(
                {
                    "representative": observation["name"],
                    "observations": [observation],
                }
            )

    records: list[dict[str, Any]] = []
    existing_codes: set[str] = set()
    for position, group in enumerate(groups, start=1):
        group_observations = group["observations"]
        name_counts: dict[str, int] = {}
        for observation in group_observations:
            name_counts[observation["name"]] = name_counts.get(observation["name"], 0) + 1
        canonical_name = sorted(
            name_counts,
            key=lambda name: (name_counts[name], len(name), name.upper()),
            reverse=True,
        )[0]
        pages = [
            observation["source_page"]
            for observation in group_observations
            if observation["source_page"] is not None
        ]
        section_pages = [
            observation["section_page"]
            for observation in group_observations
            if observation["section_page"] is not None
        ]
        variants = sorted(name_counts, key=str.upper)
        detection_keys = sorted({observation["key"] for observation in group_observations})
        average_confidence = sum(
            observation["confidence"] for observation in group_observations
        ) / max(1, len(group_observations))

        records.append(
            {
                "INCLUDE": True,
                "CODE": _generated_submachinery_code(position, existing_codes),
                "NAME": canonical_name,
                "MAKER": clean_text(main_row.get("MAKER", "")),
                "MODEL": clean_text(main_row.get("MODEL", "")),
                "TYPE": "",
                "INSTR.BOOK": clean_text(main_row.get("INSTR.BOOK", "")),
                "SPECIFICATIONS": "",
                "MCH_TP(M/S/U)": "SubMachinery",
                "FIRST PAGE": min(section_pages or pages) if (section_pages or pages) else None,
                "LAST PAGE": max(pages) if pages else None,
                "PARTS FOUND": len(group_observations),
                "CONFIDENCE": average_confidence,
                "VARIANTS": " | ".join(variants),
                "DETECTION KEYS": "|".join(detection_keys),
                "ORIGIN": "Auto-detected",
            }
        )

    frame = pd.DataFrame(records, columns=SUBMACHINERY_REVIEW_COLUMNS)
    frame["INCLUDE"] = frame["INCLUDE"].astype(bool)
    frame["FIRST PAGE"] = pd.to_numeric(frame["FIRST PAGE"], errors="coerce").astype("Int64")
    frame["LAST PAGE"] = pd.to_numeric(frame["LAST PAGE"], errors="coerce").astype("Int64")
    frame["PARTS FOUND"] = pd.to_numeric(frame["PARTS FOUND"], errors="coerce").fillna(0).astype(int)
    frame["CONFIDENCE"] = pd.to_numeric(frame["CONFIDENCE"], errors="coerce").fillna(0.70)
    return frame.sort_values(["FIRST PAGE", "NAME"], na_position="last").reset_index(drop=True)


def merge_submachinery_candidates(
    existing: pd.DataFrame | None,
    detected: pd.DataFrame | None,
) -> pd.DataFrame:
    """Merge refreshed detections while preserving user edits and manual rows."""
    if existing is None or existing.empty:
        return detected.copy() if detected is not None else empty_submachinery_review_dataframe()
    if detected is None or detected.empty:
        return existing.copy()

    result = existing.copy()
    for column in SUBMACHINERY_REVIEW_COLUMNS:
        if column not in result.columns:
            result[column] = False if column == "INCLUDE" else ""
    result = result[SUBMACHINERY_REVIEW_COLUMNS]

    for _, new_row in detected.iterrows():
        new_keys = _split_detection_keys(new_row.get("DETECTION KEYS", ""))
        match_index: Any = None
        best_score = 0.0
        for index, old_row in result.iterrows():
            old_keys = _split_detection_keys(old_row.get("DETECTION KEYS", ""))
            overlap = bool(new_keys & old_keys)
            score = _machinery_name_similarity(
                clean_text(new_row.get("NAME", "")),
                clean_text(old_row.get("NAME", "")),
            )
            if overlap or score >= 0.92:
                candidate_score = 1.0 if overlap else score
                if candidate_score > best_score:
                    best_score = candidate_score
                    match_index = index
        if match_index is None:
            result = pd.concat(
                [result, pd.DataFrame([new_row], columns=SUBMACHINERY_REVIEW_COLUMNS)],
                ignore_index=True,
            )
            continue

        # Preserve editable/user-controlled fields and refresh only evidence fields.
        for column in (
            "FIRST PAGE",
            "LAST PAGE",
            "PARTS FOUND",
            "CONFIDENCE",
            "VARIANTS",
            "DETECTION KEYS",
        ):
            result.at[match_index, column] = new_row.get(column, result.at[match_index, column])

    # Ensure generated codes remain unique and fill blanks.
    used_codes: set[str] = set()
    next_position = 1
    for index, row in result.iterrows():
        code = clean_text(row.get("CODE", ""))
        code_key = normalize_key(code)
        if not code or code_key in used_codes:
            code = _generated_submachinery_code(next_position, used_codes)
            result.at[index, "CODE"] = code
        else:
            used_codes.add(code_key)
        next_position += 1
        result.at[index, "MCH_TP(M/S/U)"] = "SubMachinery"

    return result[SUBMACHINERY_REVIEW_COLUMNS].reset_index(drop=True)


def add_manual_submachinery_candidate(
    frame: pd.DataFrame | None,
    main_row: dict[str, Any],
) -> pd.DataFrame:
    existing = frame.copy() if frame is not None else empty_submachinery_review_dataframe()
    used_codes = {
        normalize_key(value)
        for value in existing.get("CODE", pd.Series(dtype=str)).tolist()
        if clean_text(value)
    }
    row = {
        "INCLUDE": True,
        "CODE": _generated_submachinery_code(len(existing) + 1, used_codes),
        "NAME": "",
        "MAKER": clean_text(main_row.get("MAKER", "")),
        "MODEL": clean_text(main_row.get("MODEL", "")),
        "TYPE": "",
        "INSTR.BOOK": clean_text(main_row.get("INSTR.BOOK", "")),
        "SPECIFICATIONS": "",
        "MCH_TP(M/S/U)": "SubMachinery",
        "FIRST PAGE": None,
        "LAST PAGE": None,
        "PARTS FOUND": 0,
        "CONFIDENCE": 1.0,
        "VARIANTS": "",
        "DETECTION KEYS": "",
        "ORIGIN": "Manual",
    }
    new_frame = pd.DataFrame([row], columns=SUBMACHINERY_REVIEW_COLUMNS)
    if existing.empty:
        return new_frame
    return pd.concat([existing, new_frame], ignore_index=True)


def included_submachinery_rows(frame: pd.DataFrame | None) -> pd.DataFrame:
    if frame is None or frame.empty:
        return empty_additional_machinery_dataframe()
    working = frame.copy()
    if "INCLUDE" in working.columns:
        working = working[working["INCLUDE"].astype(bool)]
    for column in MACHINERY_COLUMNS:
        if column not in working.columns:
            working[column] = ""
    working["MCH_TP(M/S/U)"] = "SubMachinery"
    return working[MACHINERY_COLUMNS].reset_index(drop=True)


def apply_submachinery_assignments(
    review_frame: pd.DataFrame,
    submachinery_frame: pd.DataFrame | None,
    main_machinery: str,
    overwrite_auto_assignments: bool = True,
) -> pd.DataFrame:
    """Assign spare-part rows to approved detected sub-machineries."""
    if review_frame is None or review_frame.empty:
        return empty_review_dataframe()

    result = review_frame.copy()
    for column in REVIEW_COLUMNS:
        if column not in result.columns:
            result[column] = False if column in {"INCLUDE", "READY"} else ""

    approved = included_submachinery_rows(submachinery_frame)
    key_map: dict[str, str] = {}
    if submachinery_frame is not None and not submachinery_frame.empty:
        for _, candidate in submachinery_frame.iterrows():
            if not bool(candidate.get("INCLUDE", False)):
                continue
            target_name = clean_text(candidate.get("NAME", ""))
            if not target_name:
                continue
            keys = _split_detection_keys(candidate.get("DETECTION KEYS", ""))
            keys.add(normalize_key(target_name))
            for key in keys:
                if key:
                    key_map[key] = target_name

    approved_names = {
        normalize_key(name)
        for name in approved.get("NAME", pd.Series(dtype=str)).tolist()
        if clean_text(name)
    }
    main_key = normalize_key(main_machinery)

    for index, row in result.iterrows():
        detected = _clean_machinery_name(
            row.get("DETECTED MACHINERY", row.get("TABLE TITLE", ""))
        )
        detected_key = normalize_key(detected)
        target = key_map.get(detected_key)
        if not target and detected_key:
            best_key = ""
            best_score = 0.0
            for known_key, known_name in key_map.items():
                score = SequenceMatcher(None, detected_key, known_key).ratio()
                if score > best_score:
                    best_score = score
                    best_key = known_key
            if best_score >= 0.92:
                target = key_map.get(best_key)

        current = clean_text(row.get("MACHINERY", ""))
        current_source = clean_text(row.get("ASSIGNMENT SOURCE", ""))
        current_is_auto = (
            not current
            or normalize_key(current) == main_key
            or normalize_key(current) not in approved_names | ({main_key} if main_key else set())
            or current_source.startswith("Auto")
            or current_source.startswith("Main")
        )

        if target and (overwrite_auto_assignments or current_is_auto):
            result.at[index, "MACHINERY"] = target
            result.at[index, "ASSIGNMENT SOURCE"] = "Auto-detected sub-machinery"
        elif not current:
            result.at[index, "MACHINERY"] = clean_text(main_machinery)
            result.at[index, "ASSIGNMENT SOURCE"] = "Main machinery default"
        elif not current_source:
            result.at[index, "ASSIGNMENT SOURCE"] = "Manual assignment"

    return result[REVIEW_COLUMNS]


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
        detected_machinery = _clean_machinery_name(
            raw.get("detected_machinery", raw.get("DETECTED MACHINERY", ""))
        )
        table_title = clean_text(raw.get("table_title", raw.get("TABLE TITLE", "")))
        source_page = quantity_to_number(raw.get("source_page", raw.get("SOURCE PAGE")))
        source_page_int = int(source_page) if source_page is not None else None
        section_start = quantity_to_number(
            raw.get("section_start_page", raw.get("SECTION START PAGE"))
        )
        section_start_int = int(section_start) if section_start is not None else source_page_int
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
                "SECTION START PAGE": section_start_int,
                "TABLE TITLE": table_title,
                "CONFIDENCE": confidence,
                "DETECTED MACHINERY": detected_machinery,
                "ASSIGNMENT SOURCE": "Main machinery default",
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
    frame["SECTION START PAGE"] = pd.to_numeric(
        frame["SECTION START PAGE"], errors="coerce"
    ).astype("Int64")
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
            if "INCLUDE" in additional.columns and not bool(row.get("INCLUDE", False)):
                continue
            record = {column: clean_text(row.get(column, "")) for column in MACHINERY_COLUMNS}
            if any(record.values()):
                record["MCH_TP(M/S/U)"] = "SubMachinery"
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
    result["SECTION START PAGE"] = pd.to_numeric(
        result["SECTION START PAGE"], errors="coerce"
    ).astype("Int64")
    result["CONFIDENCE"] = pd.to_numeric(result["CONFIDENCE"], errors="coerce").fillna(0.70)
    return result[REVIEW_COLUMNS]


# ---------------------------------------------------------------------------
# Benefit template and audit workbook generation
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


def build_benefit_workbook(
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
            "The selected workbook is not the expected import template. Missing sheets: "
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
    page_classification: pd.DataFrame | None = None,
    extraction_log: Sequence[str] | None = None,
    vessels: Sequence[str] | None = None,
    submachinery_review: pd.DataFrame | None = None,
    job_metadata: dict[str, Any] | None = None,
) -> bytes:
    output = io.BytesIO()
    pages_frame = pd.DataFrame(extracted_pages, columns=["SOURCE PAGE", "OCR MARKDOWN"])
    classification_frame = (
        page_classification.copy()
        if page_classification is not None and not page_classification.empty
        else pd.DataFrame(columns=PAGE_CLASSIFICATION_COLUMNS)
    )
    log_frame = pd.DataFrame({"MESSAGE": list(extraction_log or [])})
    sub_frame = (
        submachinery_review.copy()
        if submachinery_review is not None and not submachinery_review.empty
        else empty_submachinery_review_dataframe()
    )

    summary_records: list[dict[str, str]] = [
        {
            "FIELD": "Generated at (UTC)",
            "VALUE": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"),
        },
        {
            "FIELD": "Vessels",
            "VALUE": ", ".join(clean_text(value) for value in (vessels or []) if clean_text(value)),
        },
    ]
    for key, value in (job_metadata or {}).items():
        summary_records.append({"FIELD": clean_text(key), "VALUE": clean_text(value)})
    summary_frame = pd.DataFrame(summary_records, columns=["FIELD", "VALUE"])

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary_frame.to_excel(writer, index=False, sheet_name="Job Summary")
        pages_frame.to_excel(writer, index=False, sheet_name="OCR Pages")
        classification_frame.to_excel(writer, index=False, sheet_name="Page Classification")
        machinery_frame.to_excel(writer, index=False, sheet_name="Machinery Review")
        sub_frame.to_excel(writer, index=False, sheet_name="Sub-machinery Review")
        review_frame.to_excel(writer, index=False, sheet_name="Spare Parts Review")
        log_frame.to_excel(writer, index=False, sheet_name="Extraction Log")

        workbook = writer.book
        header_format = workbook.add_format(
            {"bold": True, "bg_color": "#D9EAF7", "border": 1}
        )
        wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})

        sheet_frames = (
            ("Job Summary", summary_frame),
            ("OCR Pages", pages_frame),
            ("Page Classification", classification_frame),
            ("Machinery Review", machinery_frame),
            ("Sub-machinery Review", sub_frame),
            ("Spare Parts Review", review_frame),
            ("Extraction Log", log_frame),
        )
        for sheet_name, frame in sheet_frames:
            sheet = writer.sheets[sheet_name]
            for column_index, column_name in enumerate(frame.columns):
                sheet.write(0, column_index, column_name, header_format)
            sheet.freeze_panes(1, 0)
            if len(frame.columns):
                sheet.autofilter(0, 0, max(1, len(frame)), len(frame.columns) - 1)

        writer.sheets["Job Summary"].set_column(0, 0, 26)
        writer.sheets["Job Summary"].set_column(1, 1, 90, wrap_format)
        writer.sheets["OCR Pages"].set_column(0, 0, 12)
        writer.sheets["OCR Pages"].set_column(1, 1, 100, wrap_format)
        writer.sheets["Page Classification"].set_column(0, 3, 18)
        writer.sheets["Page Classification"].set_column(4, 4, 70, wrap_format)
        writer.sheets["Machinery Review"].set_column(0, len(MACHINERY_COLUMNS) - 1, 22)
        writer.sheets["Sub-machinery Review"].set_column(
            0, max(0, len(SUBMACHINERY_REVIEW_COLUMNS) - 1), 20
        )
        writer.sheets["Sub-machinery Review"].set_column(
            SUBMACHINERY_REVIEW_COLUMNS.index("VARIANTS"),
            SUBMACHINERY_REVIEW_COLUMNS.index("VARIANTS"),
            55,
            wrap_format,
        )
        writer.sheets["Spare Parts Review"].set_column(0, len(REVIEW_COLUMNS) - 1, 20)
        writer.sheets["Spare Parts Review"].set_column(
            REVIEW_COLUMNS.index("DESCRIPTION"),
            REVIEW_COLUMNS.index("DESCRIPTION"),
            50,
            wrap_format,
        )
        writer.sheets["Spare Parts Review"].set_column(
            REVIEW_COLUMNS.index("TABLE TITLE"),
            REVIEW_COLUMNS.index("TABLE TITLE"),
            40,
            wrap_format,
        )
        writer.sheets["Spare Parts Review"].set_column(
            REVIEW_COLUMNS.index("WARNING"),
            REVIEW_COLUMNS.index("WARNING"),
            45,
            wrap_format,
        )
        writer.sheets["Extraction Log"].set_column(0, 0, 100, wrap_format)

    return output.getvalue()


# Backward-compatible name used by app.py.
def build_workbook(
    template_bytes: bytes,
    machinery_frame: pd.DataFrame,
    review_frame: pd.DataFrame,
    clear_existing: bool = True,
) -> bytes:
    return build_benefit_workbook(
        template_bytes=template_bytes,
        machinery_frame=machinery_frame,
        review_frame=review_frame,
        clear_existing=clear_existing,
    )


def safe_filename(value: str, fallback: str = "spare_parts") -> str:
    stem = Path(clean_text(value)).stem
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._-")
    return stem or fallback
