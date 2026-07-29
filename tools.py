from __future__ import annotations

import base64
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


TOOLS_VERSION = "4.5"

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
    "SECTION CODE",
    "SECTION MAKER",
    "SECTION MODEL",
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


def _normalize_api_key(api_key: str) -> str:
    """Normalize a Mistral key without ever logging or returning it in an error."""
    key = str(api_key or "").strip().strip('"').strip("'").strip()
    if key.lower().startswith("bearer "):
        key = key[7:].strip()
    if not key:
        raise RuntimeError("A Mistral API key is required.")
    return key


def _safe_api_error_text(value: Any, api_key: str = "") -> str:
    """Return a concise error message with credentials and very long bodies removed."""
    text = str(value or "").replace("\x00", " ")
    if api_key:
        text = text.replace(api_key, "***REDACTED***")
    # Some upstream libraries include the key after labels such as api_key= or Bearer.
    text = re.sub(
        r"(?i)(api[_ -]?key\s*[:=]\s*|authorization\s*[:=]\s*bearer\s+)([^\s,;\]\[{}]+)",
        r"\1***REDACTED***",
        text,
    )
    text = re.sub(r"\s+", " ", text).strip()
    return text[:1200]


def _mistral_ocr_request(
    api_key: str,
    document: dict[str, str],
    model: str = "mistral-ocr-latest",
    timeout_seconds: int = 600,
    max_retries: int = 5,
) -> dict[str, Any]:
    """Call Mistral's official OCR REST endpoint directly.

    This intentionally bypasses ``py-mistral-helper`` and its transitive SDK
    dependencies. PDF and image bytes are sent as supported base64 data URLs,
    while external sources are passed through as ordinary URLs.
    """
    key = _normalize_api_key(api_key)
    endpoint = os.getenv("MISTRAL_OCR_URL", "https://api.mistral.ai/v1/ocr")
    payload = {
        "model": model,
        "document": document,
        "include_image_base64": False,
    }
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }

    retryable_statuses = {404, 408, 409, 425, 429, 500, 502, 503, 504}
    last_error: Exception | None = None

    for attempt in range(1, max(1, int(max_retries)) + 1):
        try:
            response = requests.post(
                endpoint,
                headers=headers,
                json=payload,
                timeout=max(30, int(timeout_seconds)),
            )

            if 200 <= response.status_code < 300:
                try:
                    body = response.json()
                except ValueError as exc:
                    raise RuntimeError("Mistral OCR returned non-JSON output.") from exc
                if not isinstance(body, dict):
                    raise RuntimeError("Mistral OCR returned an unexpected response type.")
                return body

            try:
                body_value: Any = response.json()
            except ValueError:
                body_value = response.text
            detail = _safe_api_error_text(body_value, key)

            if response.status_code in retryable_statuses and attempt < max_retries:
                retry_after = response.headers.get("Retry-After", "")
                try:
                    delay = max(1.0, float(retry_after))
                except (TypeError, ValueError):
                    delay = min(20.0, float(2 ** (attempt - 1)))
                time.sleep(delay)
                continue

            if response.status_code in {401, 403}:
                raise RuntimeError(
                    f"Mistral OCR authentication failed (HTTP {response.status_code}). "
                    "Verify the key stored in Streamlit Secrets and its workspace access. "
                    f"Service response: {detail or 'No additional details.'}"
                )
            if response.status_code == 402:
                raise RuntimeError(
                    "Mistral OCR rejected the request because billing or workspace "
                    f"payment is not enabled (HTTP 402). Service response: {detail}"
                )
            if response.status_code == 429:
                raise RuntimeError(
                    "Mistral OCR rate or usage limit reached (HTTP 429). Wait and retry, "
                    f"or check the workspace Limits and Usage pages. Service response: {detail}"
                )
            raise RuntimeError(
                f"Mistral OCR request failed (HTTP {response.status_code}). "
                f"Service response: {detail or 'No additional details.'}"
            )
        except requests.Timeout as exc:
            last_error = exc
            if attempt < max_retries:
                time.sleep(min(20.0, float(2 ** (attempt - 1))))
                continue
        except requests.RequestException as exc:
            last_error = exc
            if attempt < max_retries:
                time.sleep(min(20.0, float(2 ** (attempt - 1))))
                continue
        except RuntimeError:
            raise

    raise RuntimeError(
        "Mistral OCR could not be reached after repeated attempts: "
        + _safe_api_error_text(last_error, key)
    )


def _pdf_data_url(pdf_bytes: bytes) -> str:
    encoded = base64.b64encode(pdf_bytes).decode("ascii")
    return f"data:application/pdf;base64,{encoded}"


def _image_data_url(image_bytes: bytes, suffix: str) -> str:
    extension = str(suffix or "").lower().lstrip(".")
    mime_type = {
        "png": "image/png",
        "jpg": "image/jpeg",
        "jpeg": "image/jpeg",
        "webp": "image/webp",
        "gif": "image/gif",
    }.get(extension, "image/png")
    encoded = base64.b64encode(image_bytes).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def _response_pages(response: dict[str, Any]) -> list[dict[str, Any]]:
    pages = response.get("pages", []) if isinstance(response, dict) else []
    if pages is None:
        return []
    if not isinstance(pages, list):
        raise RuntimeError("Mistral OCR returned an invalid pages collection.")
    return [page for page in pages if isinstance(page, dict)]


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
        response = _mistral_ocr_request(
            api_key=api_key,
            document={
                "type": "document_url",
                "document_url": _pdf_data_url(buffer.getvalue()),
            },
        )

        response_pages = _response_pages(response)
        if not response_pages:
            raise RuntimeError(
                f"OCR request {chunk_number}/{len(page_chunks)} completed but returned no pages."
            )

        for local_index, page in enumerate(response_pages):
            if local_index >= len(page_chunk):
                break
            original_page = page_chunk[local_index] + 1
            extracted.append((original_page, clean_markdown(page.get("markdown", ""))))

        if progress:
            progress(
                chunk_number,
                len(page_chunks),
                f"OCR request {chunk_number}/{len(page_chunks)} complete",
            )

    extracted.sort(key=lambda item: item[0])
    return extracted


def ocr_document_url(api_key: str, document_url: str) -> list[tuple[int, str]]:
    response = _mistral_ocr_request(
        api_key=api_key,
        document={
            "type": "document_url",
            "document_url": clean_text(document_url),
        },
    )
    return [
        (index + 1, clean_markdown(page.get("markdown", "")))
        for index, page in enumerate(_response_pages(response))
    ]


def ocr_image_bytes(api_key: str, image_bytes: bytes, suffix: str) -> list[tuple[int, str]]:
    response = _mistral_ocr_request(
        api_key=api_key,
        document={
            "type": "image_url",
            "image_url": _image_data_url(image_bytes, suffix),
        },
    )
    return [
        (index + 1, clean_markdown(page.get("markdown", "")))
        for index, page in enumerate(_response_pages(response))
    ]


def ocr_image_url(api_key: str, image_url: str) -> list[tuple[int, str]]:
    response = _mistral_ocr_request(
        api_key=api_key,
        document={
            "type": "image_url",
            "image_url": clean_text(image_url),
        },
    )
    return [
        (index + 1, clean_markdown(page.get("markdown", "")))
        for index, page in enumerate(_response_pages(response))
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
spare-parts manuals. Convert OCR markdown into structured rows for the Benefit
machinery and spare-parts import template.

Return one JSON object with exactly this top-level structure:
{
  "spare_parts": [
    {
      "section_code": "",
      "section_name_english": "",
      "section_maker": "",
      "section_model": "",
      "table_title": "",
      "section_start_page": 1,
      "source_part_no": "",
      "ident_no": "",
      "item_no": "",
      "description_english": "",
      "unit": "",
      "quantity": null,
      "source_page": 1,
      "confidence": 0.0
    }
  ]
}

Benefit mapping rules:
1. Extract every genuine spare-part row. Never summarize, merge, or invent rows.
2. section_code is the drawing, table, plate, assembly, or sub-machinery code that
   governs the row. Preserve all punctuation and leading zeros. If a section has
   two printed codes, join them with " / ".
3. section_name_english is the English assembly/sub-machinery title only, in
   UPPERCASE. If the source already prints several languages, choose the English
   wording. Translate only when no English wording is printed.
4. section_maker and section_model come from the relevant PDF page/section. A
   section-specific maker/model overrides the manual-level maker/model. Do not
   invent either value.
5. The first table column is often a drawing position even when its multilingual
   header contains "Part-No.". When it contains sequential callouts such as 1,
   2, 3, 1-12, P101, etc., return it as item_no.
6. ident_no is the identifier under headers such as Ident-Nr., Ident-No., Code,
   Material Code, Spare Part Code, or equivalent. This value will populate BOTH
   Benefit CODE and Benefit PART NO.
7. source_part_no is any separate manufacturer Part No. printed by the source.
   Capture it only for audit context; it is not used automatically in Benefit.
8. description_english is the individual spare-part name only, in ENGLISH and
   UPPERCASE. When German/English/French are printed together, return only the
   English phrase. Preserve dimensions, standards, stage numbers, and symbols.
9. quantity must be numeric or null. unit is PCS, SET, or an empty string.
10. source_page is the PAGE marker supplied in the user message.
11. section_start_page is the first PDF page where the section begins, including
    its sectional/exploded drawing page when that page precedes the parts table.
12. Continuation pages keep exactly the same section_code, section_name_english,
    section_maker, section_model, and section_start_page.
13. Do not convert contents/index entries, drawing callouts without a parts table,
    page numbers, headers, or prose into spare-part rows.
14. confidence is 0 to 1 and reflects OCR quality, row alignment, English-language
    selection, and section matching.

Critical example:
Source columns:
  Teil-Nr./Part-No. | Ident-Nr./Ident-No. | Benennung/Designation | Qty
Source row:
  2 | 10.10.10.40 | Gleitlager / slide bearing / palier | 1
Return:
  item_no="2", ident_no="10.10.10.40",
  description_english="SLIDE BEARING", quantity=1.
Do NOT return the section drawing code as the spare-part identifier.
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
    catalog_hint: str = "",
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
    if catalog_hint:
        prompt += (
            "Authoritative section catalogue detected from this same PDF. Use these "
            "codes and English names whenever the current page belongs to one of them:\n"
            f"{catalog_hint}\n\n"
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
        if section_start is not None and page_number is not None:
            section_start = min(section_start, page_number)
        row["section_start_page"] = section_start or page_number

        section_name = clean_text(
            row.get("section_name_english", row.get("detected_machinery", ""))
        ).upper()
        description = clean_text(
            row.get("description_english", row.get("description", ""))
        ).upper()
        ident_no = clean_text(row.get("ident_no", row.get("code", "")))
        legacy_part_no = clean_text(row.get("part_no", ""))
        section_code = clean_text(row.get("section_code", ""))
        if not ident_no and legacy_part_no and normalize_key(legacy_part_no) != normalize_key(section_code):
            ident_no = legacy_part_no

        row["section_code"] = section_code
        row["section_name_english"] = section_name
        row["detected_machinery"] = section_name
        row["section_maker"] = clean_text(row.get("section_maker", "")).upper()
        row["section_model"] = clean_text(row.get("section_model", "")).upper()
        row["description_english"] = description
        row["description"] = description
        row["ident_no"] = ident_no
        row["code"] = ident_no
        row["part_no"] = ident_no
        row["item_no"] = clean_text(row.get("item_no", ""))

        table_title = clean_text(row.get("table_title"))
        if not section_name and table_title and not _is_generic_machinery_name(table_title):
            row["section_name_english"] = table_title.upper()
            row["detected_machinery"] = table_title.upper()
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
    catalog_hint = _catalog_prompt_hint(
        build_section_catalog(extracted_pages).get("sections", [])
    )

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
                        "content": _build_extraction_prompt(batch, additional_instructions, catalog_hint),
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
# Deterministic Benefit table parsing and section catalogue
# ---------------------------------------------------------------------------


def _split_markdown_row(line: str) -> list[str]:
    stripped = line.strip().strip("|")
    return [clean_text(cell.replace("\\|", "|").replace("**", "")) for cell in re.split(r"(?<!\\)\|", stripped)]


def _is_markdown_separator(line: str) -> bool:
    cells = _split_markdown_row(line)
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells)


def _markdown_table_blocks(markdown: str) -> list[list[list[str]]]:
    lines = str(markdown or "").splitlines()
    blocks: list[list[list[str]]] = []
    index = 0
    while index + 1 < len(lines):
        if "|" in lines[index] and _is_markdown_separator(lines[index + 1]):
            block = [_split_markdown_row(lines[index])]
            index += 2
            while index < len(lines) and "|" in lines[index]:
                block.append(_split_markdown_row(lines[index]))
                index += 1
            blocks.append(block)
        else:
            index += 1
    return blocks


def _canonical_source_header(header: str) -> str | None:
    key = re.sub(r"[^a-z0-9]+", "", clean_text(header).lower())
    if not key:
        return None
    if "ident" in key or key in {"code", "partcode", "sparepartcode", "materialcode", "articlecode"}:
        return "ident_no"
    if any(token in key for token in ("description", "designation", "benennung", "partname", "denomination")):
        return "description_raw"
    if any(token in key for token in ("quantity", "qty", "qnt", "menge", "numberoff", "nooff")):
        return "quantity"
    if any(token in key for token in ("itemno", "itemnumber", "positionno", "position", "posno", "refno", "referenceno", "indexno")):
        return "item_no"
    if any(token in key for token in ("partno", "partnumber", "teilnr", "catalogno", "orderingno", "stockno", "pno")):
        return "source_part_no"
    if key in {"unit", "uom", "unitofmeasure"}:
        return "unit"
    return None


def _find_data_header_index(rows: Sequence[Sequence[str]]) -> int | None:
    for row_index, row in enumerate(rows[:5]):
        joined = " ".join(clean_text(value).lower() for value in row)
        has_description = any(token in joined for token in ("description", "designation", "benennung", "part name"))
        has_identifier = any(token in joined for token in ("ident", "code", "part-no", "part no", "item", "position"))
        if has_description and has_identifier:
            return row_index
    return None


_HYPHEN_SECTION_CODE_RE = re.compile(
    r"\b(?:[A-Z]{1,6}-)?\d+[A-Z0-9.]*-\d+[A-Z0-9.]*-\d+[A-Z0-9.]*\b",
    flags=re.IGNORECASE,
)
_PAREN_SECTION_CODE_RE = re.compile(r"\(([A-Z]{1,6}\d{2,5})\)", flags=re.IGNORECASE)


def _section_code_tokens(value: Any, permissive: bool = False) -> list[str]:
    text = clean_text(value).upper()
    found: list[str] = []
    for match in _HYPHEN_SECTION_CODE_RE.findall(text):
        token = clean_text(match).upper()
        if token and token not in found:
            found.append(token)
    for match in _PAREN_SECTION_CODE_RE.findall(text):
        token = clean_text(match).upper()
        if token and token not in found:
            found.append(token)
    if permissive and not found:
        for token in re.findall(r"\b[A-Z]{1,5}\d{2,5}\b", text):
            # Avoid treating common model forms such as L350 as section codes unless
            # they are explicitly parenthesised or the whole cell is code-like.
            if token.startswith("L") and token[1:].isdigit() and len(token) <= 5:
                continue
            if token not in found:
                found.append(token)
    return found


def _join_section_codes(tokens: Sequence[str]) -> str:
    cleaned = [clean_text(token).upper() for token in tokens if clean_text(token)]
    return " / ".join(dict.fromkeys(cleaned))


def _english_variant(value: Any) -> str:
    text = clean_text(value)
    if not text:
        return ""
    slash_parts = [clean_text(part) for part in re.split(r"\s*/\s*", text) if clean_text(part)]
    if len(slash_parts) >= 3:
        return slash_parts[1].upper()
    line_parts = [clean_text(part) for part in re.split(r"[\r\n]+", str(value)) if clean_text(part)]
    if len(line_parts) >= 3:
        return line_parts[1].upper()
    return text.upper()


def _best_effort_english_description(value: Any) -> tuple[str, bool]:
    """Return uppercase English text and whether language review is still needed."""
    text = clean_text(value)
    if not text:
        return "", True
    if "/" in text:
        parts = [clean_text(part) for part in re.split(r"\s*/\s*", text) if clean_text(part)]
        if len(parts) >= 3:
            return parts[1].upper(), False
    lines = [clean_text(part) for part in re.split(r"[\r\n]+", str(value)) if clean_text(part)]
    if len(lines) >= 3:
        return lines[1].upper(), False
    # A single phrase from AI is already acceptable. Long concatenated multilingual
    # cells remain visible as an exception rather than being silently mistranslated.
    multilingual_markers = (
        " dés", " soupape", " joint ", " tuyau", " vis ", " arbre ", " carter",
        " und ", " schraube", " dichtung", " lager", " stufe", " zylinder",
    )
    lower = f" {text.lower()} "
    needs_review = sum(marker in lower for marker in multilingual_markers) >= 2
    return text.upper(), needs_review


def _page_metadata(page_number: int, markdown: str) -> dict[str, Any]:
    metadata: dict[str, Any] = {
        "page": int(page_number),
        "section_code": "",
        "section_name_raw": "",
        "section_name_english": "",
        "maker": "",
        "model": "",
    }
    blocks = _markdown_table_blocks(markdown)
    for block in blocks:
        header_index = _find_data_header_index(block)
        if header_index is None or header_index <= 0:
            continue
        meta_row = block[0]
        nonempty = [clean_text(value) for value in meta_row if clean_text(value)]
        if meta_row:
            metadata["maker"] = clean_text(meta_row[0]).upper()
        if len(meta_row) >= 3:
            metadata["section_name_raw"] = clean_text(meta_row[-2])
        tail = clean_text(meta_row[-1]) if meta_row else ""
        tokens = _section_code_tokens(tail, permissive=True)
        if tokens:
            metadata["section_code"] = _join_section_codes(tokens)
        model_part = re.split(r"\b(?:TAFEL|TABLE|PLANCHE|DRAWING|DWG)\b", tail, maxsplit=1, flags=re.I)[0]
        metadata["model"] = clean_text(model_part).upper()
        break

    text = str(markdown or "")
    if not metadata["section_code"]:
        tokens = _section_code_tokens(text)
        if tokens:
            metadata["section_code"] = _join_section_codes(tokens)
    if not metadata["maker"]:
        bold_candidates = re.findall(r"\*\*([^*]{2,60})\*\*", text)
        for candidate in bold_candidates:
            candidate_clean = clean_text(candidate).upper()
            if re.search(r"[A-Z]", candidate_clean) and not re.search(r"\d", candidate_clean):
                if candidate_clean not in {"DESCRIPTION", "DESIGNATION", "QUANTITY", "TABLE"}:
                    metadata["maker"] = candidate_clean
                    break
    if not metadata["section_name_raw"]:
        lines = [
            clean_text(line.lstrip("# "))
            for line in text.splitlines()
            if clean_text(line.lstrip("# ")) and not line.lstrip().startswith("![")
        ]
        stop_index = next(
            (idx for idx, line in enumerate(lines) if re.search(r"\b(?:TAFEL|TABLE|PLANCHE)\b", line, flags=re.I)),
            None,
        )
        title_lines = lines[:stop_index] if stop_index is not None else lines[:5]
        title_lines = [line for line in title_lines if not _section_code_tokens(line)]
        title_lines = [line for line in title_lines if not re.fullmatch(r"[A-Z]?\s*\d+(?:\s*-\s*[A-Z]?\s*\d+)*", line, flags=re.I)]
        if len(title_lines) >= 3:
            metadata["section_name_english"] = title_lines[1].upper()
            metadata["section_name_raw"] = " / ".join(title_lines[:3])
        elif title_lines:
            metadata["section_name_raw"] = title_lines[0]
    if not metadata["section_name_english"]:
        metadata["section_name_english"] = _english_variant(metadata["section_name_raw"])
    if not metadata["model"]:
        model_match = re.search(
            r"\b(?:COMPRESSOR|KOMPRESSOR|ENGINE|GENERATOR|PUMP)\s+([A-Z0-9][A-Z0-9 ._/-]{1,30})",
            text,
            flags=re.I,
        )
        if model_match:
            metadata["model"] = clean_text(model_match.group(1)).upper()
    return metadata


def _direct_table_rows(extracted_pages: Sequence[tuple[int, str]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for page_number, markdown in extracted_pages:
        metadata = _page_metadata(int(page_number), markdown)
        for block in _markdown_table_blocks(markdown):
            header_index = _find_data_header_index(block)
            if header_index is None:
                continue
            headers = block[header_index]
            mappings = [_canonical_source_header(header) for header in headers]
            # When Ident-No. exists, a separate Part-No. column containing drawing
            # callouts is ITEM NO by the user's Benefit mapping rule.
            if "ident_no" in mappings and "item_no" not in mappings and "source_part_no" in mappings:
                mappings[mappings.index("source_part_no")] = "item_no"
            for values in block[header_index + 1 :]:
                padded = list(values) + [""] * max(0, len(headers) - len(values))
                record: dict[str, Any] = {
                    "source_page": int(page_number),
                    "section_code": metadata.get("section_code", ""),
                    "section_name_english": metadata.get("section_name_english", ""),
                    "section_maker": metadata.get("maker", ""),
                    "section_model": metadata.get("model", ""),
                    "table_title": metadata.get("section_name_raw", ""),
                    "confidence": 0.88,
                }
                for column_index, canonical in enumerate(mappings):
                    if canonical and column_index < len(padded):
                        record[canonical] = clean_text(padded[column_index])
                description_raw = clean_text(record.get("description_raw", ""))
                ident_no = clean_text(record.get("ident_no", ""))
                item_no = clean_text(record.get("item_no", ""))
                if description_raw and (ident_no or item_no) and re.search(r"[A-Za-zÀ-ÿ]", description_raw):
                    record["quantity"] = quantity_to_number(record.get("quantity"))
                    rows.append(record)
    return rows


def _index_sections(extracted_pages: Sequence[tuple[int, str]]) -> tuple[list[dict[str, Any]], set[int]]:
    sections: list[dict[str, Any]] = []
    index_pages: set[int] = set()
    for page_number, markdown in extracted_pages:
        page_entries: list[dict[str, Any]] = []
        for block in _markdown_table_blocks(markdown):
            for row in block:
                if len(row) < 2:
                    continue
                code_tokens = _section_code_tokens(row[-1], permissive=True)
                title = clean_text(row[0])
                if not code_tokens or not re.search(r"[A-Za-zÀ-ÿ]", title):
                    continue
                page_entries.append(
                    {
                        "code": _join_section_codes(code_tokens),
                        "aliases": code_tokens,
                        "name": _english_variant(title),
                        "maker": "",
                        "model": "",
                        "pages": set(),
                    }
                )
        if len(page_entries) >= 4:
            index_pages.add(int(page_number))
            sections.extend(page_entries)
    unique: dict[str, dict[str, Any]] = {}
    for section in sections:
        key = normalize_key(section["code"])
        if key and key not in unique:
            unique[key] = section
    return list(unique.values()), index_pages


def _global_manual_maker_model(extracted_pages: Sequence[tuple[int, str]]) -> tuple[str, str]:
    maker = ""
    model = ""
    for page_number, markdown in extracted_pages[:10]:
        metadata = _page_metadata(int(page_number), markdown)
        if not maker and metadata.get("maker"):
            maker = clean_text(metadata["maker"]).upper()
        if not model and metadata.get("model"):
            model = clean_text(metadata["model"]).upper()
        if maker and model:
            break
    return maker, model


def _looks_like_spare_table_page(markdown: str) -> bool:
    text = clean_text(markdown).lower()
    has_description = any(token in text for token in ("description", "designation", "benennung", "part name"))
    has_identifier = any(token in text for token in ("ident-nr", "ident-no", "ident no", "item no", "position", "part no", "part-no", "code"))
    has_quantity = any(token in text for token in ("qty", "quantity", "menge", "qnt"))
    return has_description and has_identifier and (has_quantity or "|" in str(markdown or ""))


def select_spare_table_pages(
    extracted_pages: Sequence[tuple[int, str]],
    fallback_pages: Sequence[tuple[int, str]] | None = None,
) -> list[tuple[int, str]]:
    """Select genuine parts-table pages while keeping OCR of all pages for section context."""
    _, index_pages = _index_sections(extracted_pages)
    selected = [
        item
        for item in extracted_pages
        if int(item[0]) not in index_pages and _looks_like_spare_table_page(item[1])
    ]
    if selected:
        return selected
    return list(fallback_pages if fallback_pages is not None else extracted_pages)


def build_section_catalog(
    extracted_pages: Sequence[tuple[int, str]],
    extracted_rows: Sequence[dict[str, Any]] | None = None,
    main_row: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Build a deterministic section catalogue keyed by source drawing/table code."""
    sections, index_pages = _index_sections(extracted_pages)
    global_maker, global_model = _global_manual_maker_model(extracted_pages)
    main_row = main_row or {}
    by_alias: dict[str, dict[str, Any]] = {}

    def register(section: dict[str, Any]) -> dict[str, Any]:
        aliases = [clean_text(value).upper() for value in section.get("aliases", []) if clean_text(value)]
        code = clean_text(section.get("code", "")) or _join_section_codes(aliases)
        name = clean_text(section.get("name", "")).upper()
        existing = next((by_alias.get(normalize_key(alias)) for alias in aliases if by_alias.get(normalize_key(alias))), None)
        if existing is None:
            existing = {
                "code": code,
                "aliases": aliases or ([code] if code else []),
                "name": name,
                "maker": clean_text(section.get("maker", "")).upper(),
                "model": clean_text(section.get("model", "")).upper(),
                "pages": set(section.get("pages", set())),
            }
            sections.append(existing) if existing not in sections else None
        else:
            if name and (not existing.get("name") or len(name) < len(existing.get("name", ""))):
                existing["name"] = name
            if clean_text(section.get("maker", "")):
                existing["maker"] = clean_text(section.get("maker", "")).upper()
            if clean_text(section.get("model", "")):
                existing["model"] = clean_text(section.get("model", "")).upper()
            existing["pages"].update(section.get("pages", set()))
            for alias in aliases:
                if alias not in existing["aliases"]:
                    existing["aliases"].append(alias)
        for alias in existing.get("aliases", []):
            by_alias[normalize_key(alias)] = existing
        return existing

    # Re-register index entries in a clean map.
    original_sections = list(sections)
    sections = []
    for section in original_sections:
        register(section)

    page_metadata: dict[int, dict[str, Any]] = {}
    for page_number, markdown in extracted_pages:
        metadata = _page_metadata(int(page_number), markdown)
        page_metadata[int(page_number)] = metadata
        if int(page_number) in index_pages:
            continue
        code_tokens = _section_code_tokens(metadata.get("section_code", ""), permissive=True)
        if code_tokens:
            section = register(
                {
                    "code": _join_section_codes(code_tokens),
                    "aliases": code_tokens,
                    "name": metadata.get("section_name_english", ""),
                    "maker": metadata.get("maker", ""),
                    "model": metadata.get("model", ""),
                    "pages": {int(page_number)},
                }
            )
            section["pages"].add(int(page_number))

    for row in extracted_rows or []:
        code_tokens = _section_code_tokens(row.get("section_code", ""), permissive=True)
        if not code_tokens:
            continue
        source_page = quantity_to_number(row.get("source_page"))
        register(
            {
                "code": _join_section_codes(code_tokens),
                "aliases": code_tokens,
                "name": clean_text(row.get("section_name_english", row.get("detected_machinery", ""))).upper(),
                "maker": clean_text(row.get("section_maker", "")).upper(),
                "model": clean_text(row.get("section_model", "")).upper(),
                "pages": {int(source_page)} if source_page is not None else set(),
            }
        )

    # Exact code appearances include drawing pages and are the authoritative start.
    page_text_lookup = {int(page): clean_text(markdown).upper() for page, markdown in extracted_pages}
    for page, text in page_text_lookup.items():
        if page in index_pages:
            continue
        for section in sections:
            if any(alias and alias in text for alias in section.get("aliases", [])):
                section["pages"].add(page)

    parts_pages = {
        int(page)
        for page, markdown in extracted_pages
        if int(page) not in index_pages and _looks_like_spare_table_page(markdown)
    }
    page_map: dict[int, dict[str, Any]] = {}
    active: dict[str, Any] | None = None
    for page in sorted(page_text_lookup):
        matches = []
        text = page_text_lookup[page]
        for section in sections:
            if any(alias and alias in text for alias in section.get("aliases", [])):
                matches.append(section)
        unique_matches = []
        seen_ids = set()
        for match in matches:
            if id(match) not in seen_ids:
                unique_matches.append(match)
                seen_ids.add(id(match))
        if len(unique_matches) == 1:
            active = unique_matches[0]
        if page in parts_pages and active is not None:
            page_map[page] = active
            active["pages"].add(page)

    fallback_maker = global_maker or clean_text(main_row.get("MAKER", "")).upper()
    fallback_model = global_model or clean_text(main_row.get("MODEL", "")).upper()
    final_sections: list[dict[str, Any]] = []
    seen_section_ids: set[int] = set()
    for section in sections:
        if id(section) in seen_section_ids:
            continue
        seen_section_ids.add(id(section))
        pages = sorted(int(page) for page in section.get("pages", set()) if int(page) not in index_pages)
        section["name"] = clean_text(section.get("name", "")).upper()
        section["maker"] = clean_text(section.get("maker", "")).upper() or fallback_maker
        section["model"] = clean_text(section.get("model", "")).upper() or fallback_model
        section["first_page"] = min(pages) if pages else None
        section["last_page"] = max(pages) if pages else None
        section["pages"] = pages
        if section.get("code") or section.get("name"):
            final_sections.append(section)
    final_sections.sort(key=lambda item: (item.get("first_page") is None, item.get("first_page") or 10**9, item.get("name", "")))
    return {
        "sections": final_sections,
        "page_map": page_map,
        "index_pages": index_pages,
        "parts_pages": parts_pages,
    }


def _catalog_prompt_hint(sections: Sequence[dict[str, Any]]) -> str:
    lines: list[str] = []
    for section in sections[:120]:
        code = clean_text(section.get("code", ""))
        name = clean_text(section.get("name", ""))
        maker = clean_text(section.get("maker", ""))
        model = clean_text(section.get("model", ""))
        if code or name:
            lines.append(f"- CODE={code}; NAME={name}; MAKER={maker}; MODEL={model}")
    return "\n".join(lines)


def _normalized_ai_row(row: dict[str, Any]) -> dict[str, Any]:
    item = dict(row)
    ident = clean_text(item.get("ident_no", item.get("code", item.get("part_no", ""))))
    section_code = clean_text(item.get("section_code", ""))
    if ident and section_code and normalize_key(ident) == normalize_key(section_code):
        ident = ""
    description, language_review = _best_effort_english_description(
        item.get("description_english", item.get("description", ""))
    )
    return {
        **item,
        "ident_no": ident,
        "item_no": clean_text(item.get("item_no", "")),
        "description_english": description,
        "language_review": bool(language_review),
        "section_code": section_code,
        "section_name_english": clean_text(item.get("section_name_english", item.get("detected_machinery", ""))).upper(),
        "section_maker": clean_text(item.get("section_maker", "")).upper(),
        "section_model": clean_text(item.get("section_model", "")).upper(),
        "source_page": int(quantity_to_number(item.get("source_page"))) if quantity_to_number(item.get("source_page")) is not None else None,
        "section_start_page": int(quantity_to_number(item.get("section_start_page"))) if quantity_to_number(item.get("section_start_page")) is not None else None,
        "quantity": quantity_to_number(item.get("quantity", item.get("QNT"))),
        "confidence": clamp_confidence(item.get("confidence", 0.70)),
    }


def prepare_benefit_rows(
    ai_rows: Sequence[dict[str, Any]],
    extracted_pages: Sequence[tuple[int, str]],
    source_document_name: str,
    main_row: dict[str, Any],
    default_unit: str = "PCS",
    catalog_pages: Sequence[tuple[int, str]] | None = None,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Repair and normalize extraction using exact PDF table structure and section codes."""
    normalized_ai = [_normalized_ai_row(row) for row in ai_rows if isinstance(row, dict)]
    direct_rows = _direct_table_rows(extracted_pages)
    section_context_pages = list(catalog_pages) if catalog_pages is not None else list(extracted_pages)
    catalog = build_section_catalog(section_context_pages, normalized_ai, main_row)
    sections = catalog["sections"]
    page_map = catalog["page_map"]
    index_pages = catalog["index_pages"]
    parts_pages = catalog["parts_pages"]

    ai_by_page_item: dict[tuple[int | None, str], dict[str, Any]] = {}
    ai_by_page_ident: dict[tuple[int | None, str], dict[str, Any]] = {}
    for row in normalized_ai:
        page = row.get("source_page")
        item_key = normalize_key(row.get("item_no", ""))
        ident_key = normalize_key(row.get("ident_no", ""))
        if item_key:
            ai_by_page_item[(page, item_key)] = row
        if ident_key:
            ai_by_page_ident[(page, ident_key)] = row

    output: list[dict[str, Any]] = []
    matched_ai_ids: set[int] = set()
    language_exceptions = 0

    def section_for(page: int | None, row: dict[str, Any]) -> dict[str, Any] | None:
        code_tokens = _section_code_tokens(row.get("section_code", ""), permissive=True)
        if code_tokens:
            code_keys = {normalize_key(token) for token in code_tokens}
            for section in sections:
                if code_keys & {normalize_key(alias) for alias in section.get("aliases", [])}:
                    return section
        if page is not None and page in page_map:
            return page_map[page]
        name_key = normalize_key(row.get("section_name_english", ""))
        if name_key:
            scored = sorted(
                ((SequenceMatcher(None, name_key, normalize_key(section.get("name", ""))).ratio(), section) for section in sections),
                key=lambda pair: pair[0],
                reverse=True,
            )
            if scored and scored[0][0] >= 0.90:
                return scored[0][1]
        return None

    if direct_rows:
        for direct in direct_rows:
            page = int(direct["source_page"])
            item_key = normalize_key(direct.get("item_no", ""))
            ident_key = normalize_key(direct.get("ident_no", ""))
            ai = ai_by_page_item.get((page, item_key)) or ai_by_page_ident.get((page, ident_key))
            if ai is not None:
                matched_ai_ids.add(id(ai))
            description_source = ""
            if ai is not None and not bool(ai.get("local_fallback", False)):
                description_source = clean_text(ai.get("description_english", ""))
            description, language_review = _best_effort_english_description(description_source)
            if not description:
                description, isolated_from_layout = _best_effort_english_description(
                    direct.get("description_raw", "")
                )
                # A flattened multilingual table cell cannot be trusted without the
                # AI English selection. Slash/line-separated cells may still be exact.
                raw_text = str(direct.get("description_raw", ""))
                has_separable_layout = "/" in raw_text or "\n" in raw_text or "\r" in raw_text
                language_review = bool(isolated_from_layout or not has_separable_layout)
            if language_review:
                language_exceptions += 1
            section = section_for(page, ai or direct)
            section_code = clean_text((section or {}).get("code", "")) or clean_text((ai or {}).get("section_code", direct.get("section_code", "")))
            section_name = clean_text((section or {}).get("name", "")) or clean_text((ai or {}).get("section_name_english", direct.get("section_name_english", "")))
            section_maker = clean_text((ai or {}).get("section_maker", "")) or clean_text((section or {}).get("maker", "")) or clean_text(direct.get("section_maker", ""))
            section_model = clean_text((ai or {}).get("section_model", "")) or clean_text((section or {}).get("model", "")) or clean_text(direct.get("section_model", ""))
            first_page = (section or {}).get("first_page") or page
            confidence = max(0.90, clamp_confidence((ai or {}).get("confidence", direct.get("confidence", 0.88))))
            if language_review:
                confidence = min(confidence, 0.60)
            ident_no = clean_text(direct.get("ident_no", ""))
            output.append(
                {
                    "source_page": page,
                    "section_start_page": int(first_page),
                    "section_code": section_code,
                    "section_name_english": section_name.upper(),
                    "detected_machinery": section_name.upper(),
                    "section_maker": section_maker.upper(),
                    "section_model": section_model.upper(),
                    "table_title": clean_text(direct.get("table_title", section_name)),
                    "ident_no": ident_no,
                    "part_no": ident_no,
                    "code": ident_no,
                    "item_no": clean_text(direct.get("item_no", "")),
                    "description_english": description.upper(),
                    "description": description.upper(),
                    "unit": normalize_unit((ai or {}).get("unit", ""), default_unit),
                    "quantity": quantity_to_number(direct.get("quantity")),
                    "confidence": confidence,
                    "language_review": language_review,
                }
            )
    else:
        # Irregular/non-Markdown manuals still use the strict AI schema.
        for ai in normalized_ai:
            page = ai.get("source_page")
            if page in index_pages:
                continue
            section = section_for(page, ai)
            ident_no = clean_text(ai.get("ident_no", ""))
            description = clean_text(ai.get("description_english", "")).upper()
            if not description or not (ident_no or clean_text(ai.get("item_no", ""))):
                continue
            output.append(
                {
                    **ai,
                    "section_code": clean_text((section or {}).get("code", ai.get("section_code", ""))),
                    "section_name_english": clean_text((section or {}).get("name", ai.get("section_name_english", ""))).upper(),
                    "detected_machinery": clean_text((section or {}).get("name", ai.get("section_name_english", ""))).upper(),
                    "section_maker": clean_text(ai.get("section_maker", "")) or clean_text((section or {}).get("maker", "")),
                    "section_model": clean_text(ai.get("section_model", "")) or clean_text((section or {}).get("model", "")),
                    "section_start_page": (section or {}).get("first_page") or ai.get("section_start_page") or page,
                    "part_no": ident_no,
                    "code": ident_no,
                    "description": description,
                }
            )

    # AI-only rows are accepted only on pages that contain an actual parts table.
    if direct_rows:
        existing_keys = {
            (int(row.get("source_page") or 0), normalize_key(row.get("item_no", "")), normalize_key(row.get("ident_no", "")))
            for row in output
        }
        for ai in normalized_ai:
            if id(ai) in matched_ai_ids or ai.get("source_page") not in parts_pages:
                continue
            key = (int(ai.get("source_page") or 0), normalize_key(ai.get("item_no", "")), normalize_key(ai.get("ident_no", "")))
            if key in existing_keys:
                continue
            if clean_text(ai.get("description_english", "")) and (clean_text(ai.get("ident_no", "")) or clean_text(ai.get("item_no", ""))):
                section = section_for(ai.get("source_page"), ai)
                ident_no = clean_text(ai.get("ident_no", ""))
                output.append(
                    {
                        **ai,
                        "section_code": clean_text((section or {}).get("code", ai.get("section_code", ""))),
                        "section_name_english": clean_text((section or {}).get("name", ai.get("section_name_english", ""))).upper(),
                        "detected_machinery": clean_text((section or {}).get("name", ai.get("section_name_english", ""))).upper(),
                        "section_maker": clean_text(ai.get("section_maker", "")) or clean_text((section or {}).get("maker", "")),
                        "section_model": clean_text(ai.get("section_model", "")) or clean_text((section or {}).get("model", "")),
                        "section_start_page": (section or {}).get("first_page") or ai.get("source_page"),
                        "part_no": ident_no,
                        "code": ident_no,
                        "description": clean_text(ai.get("description_english", "")).upper(),
                    }
                )

    # Deterministic de-duplication by source page + item + identifier.
    deduplicated: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str]] = set()
    for row in output:
        key = (
            str(row.get("source_page", "")),
            normalize_key(row.get("item_no", "")),
            normalize_key(row.get("ident_no", row.get("code", ""))),
        )
        if key in seen:
            continue
        seen.add(key)
        deduplicated.append(row)

    messages = [
        f"Deterministic table verification found {len(direct_rows)} spare-part row(s).",
        f"Detected {len(sections)} source-coded sub-machinery section(s).",
    ]
    if language_exceptions:
        messages.append(
            f"{language_exceptions} description(s) could not be confidently isolated as English and remain in the exception review queue."
        )
    return deduplicated, messages


# Backward-compatible local parser used by the app when AI extraction is disabled.
def extract_spare_parts_from_markdown_tables(
    extracted_pages: Sequence[tuple[int, str]],
) -> list[dict[str, Any]]:
    rows, _ = prepare_benefit_rows(
        ai_rows=[],
        extracted_pages=extracted_pages,
        source_document_name="",
        main_row={},
        default_unit="PCS",
    )
    for row in rows:
        row["local_fallback"] = True
    return rows


# ---------------------------------------------------------------------------
# Automatic sub-machinery detection and assignment
# ---------------------------------------------------------------------------


_GENERIC_MACHINERY_NAMES = {
    "SPARE PARTS", "SPARE PARTS LIST", "PARTS LIST", "LIST OF PARTS",
    "PARTS CATALOGUE", "PARTS CATALOG", "CATALOGUE", "CATALOG",
    "DESCRIPTION", "DESIGNATION", "ITEM NO", "ITEM NUMBER", "PART NO",
    "PART NUMBER", "QUANTITY", "DRAWING", "DRAWING NO", "TABLE",
    "CONTINUED", "LIST OF SPARE PARTS",
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
    return clean_text(text).strip(" -:;|/").upper()


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
    token_score = len(left_tokens & right_tokens) / max(1, len(left_tokens | right_tokens))
    generic_suffixes = {"ASSEMBLY", "ASSY", "UNIT", "COMPLETE"}
    containment = 0.0
    if min(len(left_key), len(right_key)) >= 8 and (left_key in right_key or right_key in left_key):
        extra_tokens = (left_tokens | right_tokens) - (left_tokens & right_tokens)
        if extra_tokens and extra_tokens <= generic_suffixes:
            containment = 0.94
    return max(sequence, token_score, containment)


def _split_detection_keys(value: Any) -> set[str]:
    return {key.strip() for key in clean_text(value).split("|") if key.strip()}


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
    source_document_name: str = "",
) -> pd.DataFrame:
    """Create source-code-driven Benefit sub-machinery rows from extracted parts."""
    if review_frame is None or review_frame.empty:
        return empty_submachinery_review_dataframe()

    observations: list[dict[str, Any]] = []
    for _, row in review_frame.iterrows():
        name = _clean_machinery_name(row.get("DETECTED MACHINERY", row.get("TABLE TITLE", "")))
        code = clean_text(row.get("SECTION CODE", "")).upper()
        if _is_generic_machinery_name(name) and not code:
            continue
        source = quantity_to_number(row.get("SOURCE PAGE"))
        section = quantity_to_number(row.get("SECTION START PAGE"))
        observations.append(
            {
                "name": name,
                "code": code,
                "maker": clean_text(row.get("SECTION MAKER", "")).upper(),
                "model": clean_text(row.get("SECTION MODEL", "")).upper(),
                "source_page": int(source) if source is not None else None,
                "section_page": int(section) if section is not None else None,
                "confidence": clamp_confidence(row.get("CONFIDENCE", 0.70)),
            }
        )
    if not observations:
        return empty_submachinery_review_dataframe()

    groups: dict[str, list[dict[str, Any]]] = {}
    for observation in observations:
        group_key = normalize_key(observation["code"]) or normalize_key(observation["name"])
        if not group_key:
            continue
        groups.setdefault(group_key, []).append(observation)

    preliminary: list[dict[str, Any]] = []
    for group_observations in groups.values():
        code_values = [obs["code"] for obs in group_observations if obs["code"]]
        name_values = [obs["name"] for obs in group_observations if obs["name"]]
        maker_values = [obs["maker"] for obs in group_observations if obs["maker"]]
        model_values = [obs["model"] for obs in group_observations if obs["model"]]
        code = max(set(code_values), key=code_values.count) if code_values else ""
        name = max(set(name_values), key=lambda value: (name_values.count(value), -len(value))) if name_values else ""
        maker = max(set(maker_values), key=maker_values.count) if maker_values else clean_text(main_row.get("MAKER", "")).upper()
        model = max(set(model_values), key=model_values.count) if model_values else clean_text(main_row.get("MODEL", "")).upper()
        pages = [obs["source_page"] for obs in group_observations if obs["source_page"] is not None]
        section_pages = [obs["section_page"] for obs in group_observations if obs["section_page"] is not None]
        first_page = min(section_pages or pages) if (section_pages or pages) else None
        preliminary.append(
            {
                "code": code,
                "name": name.upper(),
                "maker": maker,
                "model": model,
                "first_page": first_page,
                "last_page": max(pages) if pages else first_page,
                "parts": len(group_observations),
                "confidence": sum(obs["confidence"] for obs in group_observations) / max(1, len(group_observations)),
                "detection_keys": "|".join(sorted({normalize_key(obs["name"]) for obs in group_observations if obs["name"]} | {normalize_key(code)})),
                "variants": " | ".join(sorted(set(name_values), key=str.upper)),
            }
        )

    records: list[dict[str, Any]] = []
    for item in preliminary:
        benefit_name = item["name"]
        # Benefit links sheet 2 to sheet 1 through the exact machinery NAME. Keep
        # the source code visibly embedded in every automated sub-machinery name,
        # matching the proven reference-workbook convention (e.g. CYLINDER HEAD
        # (P101)) and eliminating ambiguous same-name sections.
        if benefit_name and item["code"] and f"({item['code']})" not in benefit_name:
            benefit_name = f"{benefit_name} ({item['code']})"
        records.append(
            {
                "INCLUDE": True,
                "CODE": item["code"],
                "NAME": benefit_name,
                "MAKER": item["maker"],
                "MODEL": item["model"],
                "TYPE": "",
                "INSTR.BOOK": f"PDF PAGE.{item['first_page']}" if item["first_page"] is not None else "",
                "SPECIFICATIONS": f"PDF FILE: {clean_text(source_document_name)}" if clean_text(source_document_name) else "",
                "MCH_TP(M/S/U)": "SubMachinery",
                "FIRST PAGE": item["first_page"],
                "LAST PAGE": item["last_page"],
                "PARTS FOUND": item["parts"],
                "CONFIDENCE": item["confidence"],
                "VARIANTS": item["variants"],
                "DETECTION KEYS": item["detection_keys"],
                "ORIGIN": "Auto-matched by source code",
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
    """Merge refreshed detections while preserving manually created/edited rows."""
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
        new_code = normalize_key(new_row.get("CODE", ""))
        new_keys = _split_detection_keys(new_row.get("DETECTION KEYS", ""))
        match_index: Any = None
        for index, old_row in result.iterrows():
            old_code = normalize_key(old_row.get("CODE", ""))
            old_keys = _split_detection_keys(old_row.get("DETECTION KEYS", ""))
            if (new_code and old_code == new_code) or bool(new_keys & old_keys):
                match_index = index
                break
        if match_index is None:
            result = pd.concat([result, pd.DataFrame([new_row], columns=SUBMACHINERY_REVIEW_COLUMNS)], ignore_index=True)
            continue

        old_origin = clean_text(result.at[match_index, "ORIGIN"])
        # Auto rows are fully refreshed. Manual rows retain all user-controlled values.
        refresh_columns = (
            "FIRST PAGE", "LAST PAGE", "PARTS FOUND", "CONFIDENCE", "VARIANTS", "DETECTION KEYS"
        )
        if old_origin.startswith("Auto"):
            refresh_columns = (
                "CODE", "NAME", "MAKER", "MODEL", "TYPE", "INSTR.BOOK", "SPECIFICATIONS",
                "FIRST PAGE", "LAST PAGE", "PARTS FOUND", "CONFIDENCE", "VARIANTS", "DETECTION KEYS", "ORIGIN",
            )
        for column in refresh_columns:
            result.at[match_index, column] = new_row.get(column, result.at[match_index, column])

    result["MCH_TP(M/S/U)"] = "SubMachinery"
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
    """Assign spare rows by exact source section code, then by canonical name."""
    if review_frame is None or review_frame.empty:
        return empty_review_dataframe()
    result = review_frame.copy()
    for column in REVIEW_COLUMNS:
        if column not in result.columns:
            result[column] = False if column in {"INCLUDE", "READY"} else ""

    code_map: dict[str, str] = {}
    name_map: dict[str, str] = {}
    approved_names: set[str] = set()
    if submachinery_frame is not None and not submachinery_frame.empty:
        for _, candidate in submachinery_frame.iterrows():
            if not bool(candidate.get("INCLUDE", False)):
                continue
            target_name = clean_text(candidate.get("NAME", ""))
            if not target_name:
                continue
            approved_names.add(normalize_key(target_name))
            code_key = normalize_key(candidate.get("CODE", ""))
            if code_key:
                code_map[code_key] = target_name
            for key in _split_detection_keys(candidate.get("DETECTION KEYS", "")) | {normalize_key(target_name)}:
                if key:
                    name_map[key] = target_name

    main_key = normalize_key(main_machinery)
    for index, row in result.iterrows():
        section_code_key = normalize_key(row.get("SECTION CODE", ""))
        detected_key = normalize_key(row.get("DETECTED MACHINERY", row.get("TABLE TITLE", "")))
        target = code_map.get(section_code_key) or name_map.get(detected_key)
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
            result.at[index, "ASSIGNMENT SOURCE"] = "Auto-matched by section code"
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
        identifier = clean_text(raw.get("ident_no", raw.get("code", raw.get("CODE", raw.get("part_no", raw.get("PART NO", ""))))))
        description = clean_text(raw.get("description_english", raw.get("description", raw.get("DESCRIPTION", "")))).upper()
        item_no = clean_text(raw.get("item_no", raw.get("ITEM NO", "")))
        detected_machinery = _clean_machinery_name(raw.get("section_name_english", raw.get("detected_machinery", raw.get("DETECTED MACHINERY", ""))))
        section_code = clean_text(raw.get("section_code", raw.get("SECTION CODE", ""))).upper()
        section_maker = clean_text(raw.get("section_maker", raw.get("SECTION MAKER", ""))).upper()
        section_model = clean_text(raw.get("section_model", raw.get("SECTION MODEL", ""))).upper()
        table_title = clean_text(raw.get("table_title", raw.get("TABLE TITLE", detected_machinery)))
        source_page = quantity_to_number(raw.get("source_page", raw.get("SOURCE PAGE")))
        source_page_int = int(source_page) if source_page is not None else None
        section_start = quantity_to_number(raw.get("section_start_page", raw.get("SECTION START PAGE")))
        section_start_int = int(section_start) if section_start is not None else source_page_int
        quantity = quantity_to_number(raw.get("quantity", raw.get("QNT")))
        unit = normalize_unit(raw.get("unit", raw.get("UNIT", "")), default_unit)
        confidence = clamp_confidence(raw.get("confidence", raw.get("CONFIDENCE", 0.70)))
        duplicate_key = (str(source_page_int or ""), normalize_key(identifier), normalize_key(item_no), normalize_key(description))
        if duplicate_key in seen:
            continue
        seen.add(duplicate_key)
        useful = bool(description and (identifier or item_no))
        normalized.append(
            {
                "INCLUDE": useful,
                "READY": False,
                "MACHINERY": clean_text(default_machinery),
                "PART NO": identifier,
                "DESCRIPTION": description,
                "CODE": identifier,
                "ITEM NO": item_no,
                "UNIT": unit,
                "QNT": quantity,
                "SOURCE PAGE": source_page_int,
                "SECTION START PAGE": section_start_int,
                "TABLE TITLE": table_title,
                "SECTION CODE": section_code,
                "SECTION MAKER": section_maker,
                "SECTION MODEL": section_model,
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
    frame["SECTION START PAGE"] = pd.to_numeric(frame["SECTION START PAGE"], errors="coerce").astype("Int64")
    frame["CONFIDENCE"] = pd.to_numeric(frame["CONFIDENCE"], errors="coerce").fillna(0.70)
    return frame


def merge_review_dataframes(existing: pd.DataFrame, new: pd.DataFrame) -> pd.DataFrame:
    if existing is None or existing.empty:
        return new.copy()
    if new is None or new.empty:
        return existing.copy()

    combined = pd.concat([existing, new], ignore_index=True)
    def merge_key(row: pd.Series) -> str:
        source_page = str(row.get("SOURCE PAGE", ""))
        section_code = normalize_key(row.get("SECTION CODE", ""))
        item_no = normalize_key(row.get("ITEM NO", ""))
        # A manual PART NO override must not create a duplicate when the same
        # page range is appended later. Source page + section + drawing position
        # is the stable row identity whenever a position exists.
        if item_no:
            return "|".join([source_page, section_code, "ITEM", item_no])
        return "|".join(
            [
                source_page,
                section_code,
                "IDENT",
                normalize_key(row.get("CODE", row.get("PART NO", ""))),
                normalize_key(row.get("DESCRIPTION", "")),
            ]
        )

    keys = combined.apply(merge_key, axis=1)
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
                record["NAME"] = clean_text(record.get("NAME", "")).upper()
                record["MAKER"] = clean_text(record.get("MAKER", "")).upper()
                record["MODEL"] = clean_text(record.get("MODEL", "")).upper()
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

    # Benefit automation rules: descriptions are uppercase and Ident/Code is copied
    # to both CODE and PART NO unless a reviewer intentionally changed PART NO.
    result["DESCRIPTION"] = result["DESCRIPTION"].map(lambda value: clean_text(value).upper())
    for index, row in result.iterrows():
        code = clean_text(row.get("CODE", ""))
        part_no = clean_text(row.get("PART NO", ""))
        if code and not part_no:
            result.at[index, "PART NO"] = code
        elif part_no and not code:
            result.at[index, "CODE"] = part_no

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
            row_messages.append("Missing sub-machinery")
            blocking = True
        elif normalize_key(machinery) not in valid_names:
            row_messages.append("Sub-machinery is not present on sheet 1")
            blocking = True
        if not description:
            row_messages.append("Missing English description")
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
            normalize_key(machinery), normalize_key(part_no), normalize_key(item_no), normalize_key(description)
        )
        if duplicate_counter.get(duplicate_key, 0) > 1:
            row_messages.append("Possible duplicate")
            if not allow_duplicates:
                blocking = True
        confidence = clamp_confidence(row.get("CONFIDENCE", 0.70))
        # OCR confidence is advisory only. A user-corrected row must become READY
        # when all required Benefit fields and machinery links are valid.
        if confidence < 0.65:
            row_messages.append("Very low OCR confidence - manually verified fields recommended")
        elif confidence < 0.80:
            row_messages.append("Low OCR confidence - verify when practical")
        warnings.append("; ".join(row_messages))
        ready_values.append(not blocking)

    result["WARNING"] = warnings
    result["READY"] = ready_values
    result["QNT"] = pd.to_numeric(result["QNT"], errors="coerce")
    result["SOURCE PAGE"] = pd.to_numeric(result["SOURCE PAGE"], errors="coerce").astype("Int64")
    result["SECTION START PAGE"] = pd.to_numeric(result["SECTION START PAGE"], errors="coerce").astype("Int64")
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


def _xlsx_content_width(frame: pd.DataFrame, column: str, minimum: int = 10, maximum: int = 60) -> int:
    if column not in frame.columns:
        return minimum
    values = [clean_text(column)] + [clean_text(value) for value in frame[column].head(1000).tolist()]
    longest = max((len(value) for value in values), default=minimum)
    return max(minimum, min(maximum, int(longest * 1.05) + 2))


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
                maximum = 100 if column_name in {"OCR MARKDOWN", "MESSAGE"} else 60
                width = _xlsx_content_width(frame, column_name, minimum=10, maximum=maximum)
                sheet.set_column(column_index, column_index, width, wrap_format if width >= 35 else None)
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
