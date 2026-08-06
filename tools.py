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


TOOLS_VERSION = "4.9.3"

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


MULTILINGUAL_ORDER_CATALOG_PROFILE = "multilingual_order_catalog"


def _is_multilingual_order_catalog(markdown: Any) -> bool:
    """Recognize older multilingual catalogues whose real identifier is Order-No.

    These catalogues commonly place a drawing position at the far left, followed
    by burner/equipment applicability, size, Order-No., weight, and three parallel
    German/English/French designation columns.  The layout is materially different
    from the Ident-No. tables used by newer manuals, so it needs its own mapping.
    """
    text = clean_text(markdown).lower()
    order_signal = any(
        token in text
        for token in (
            "order-no",
            "order no",
            "orderno",
            "bestell-nr",
            "bestell nr",
            "no de commande",
        )
    )
    multilingual_signal = (
        "designation" in text
        and any(token in text for token in ("bezeichnung", "benennung"))
    )
    layout_signal = any(
        token in text
        for token in (
            "burner serie",
            "burner series",
            "brenner-typenreihe",
            "type bruleur",
            "type brûleur",
            "pict.",
            "photo",
            "appr. kg",
        )
    )
    return bool(order_signal and multilingual_signal and layout_signal)


def _document_extraction_profile(
    extracted_pages: Sequence[tuple[int, str]],
) -> str:
    sample = "\n".join(str(markdown or "") for _, markdown in extracted_pages[:12])
    if _is_multilingual_order_catalog(sample):
        return MULTILINGUAL_ORDER_CATALOG_PROFILE
    return ""


def _profile_prompt(profile: str) -> str:
    if profile != MULTILINGUAL_ORDER_CATALOG_PROFILE:
        return ""
    return """
AUTOMATIC LAYOUT PROFILE - MULTILINGUAL ORDER-NO. CATALOGUE:
- The column headed Bestell-Nr. / Order-No. / No de commande is the genuine
  spare-part identifier. Copy it to ident_no. It will populate both PART NO and CODE.
- Bild / Pict. / Photo is the drawing position and must populate item_no. It is
  never the spare-part code.
- Use only the middle English Designation column. Ignore the German Bezeichnung
  and the final French Designation column.
- Burner serie / Type bruleur and Size/Grandeur describe applicability. They are
  not part numbers, quantities, or sub-machinery codes.
- ca. kg / appr. kg / env. kg is WEIGHT. Never map it to quantity. Return quantity=null.
- A single printed item position can contain several burner-series/size variants
  and several Order-No. values. Return one row for EVERY Order-No.; repeat the same
  item_no and English description on each variant row.
- German assembly words such as Brennermotor, Luftregelung, Stellantrieb, and
  Brennergehause are hierarchy/title text, not a manufacturer. Leave section_maker
  empty unless the page explicitly labels a maker/manufacturer or shows a clear brand.
- Section headings are simple numbered English headings such as "3. Servo drive
  for oil burners". Use the integer as section_code and the exact English heading
  as section_name_english, without appending the code in parentheses. Carry it
  through consecutive continuation pages until the next numbered section heading.
- A repeated item position, Order-No., or individual spare description is never a
  hierarchy heading. Do not promote table-body rows to sub-machineries.
""".strip()


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
   invent either value. Only use an explicit maker/manufacturer label or a clear
   brand/logo. German assembly/title words such as Brennermotor, Luftregelung,
   Stellantrieb, Brennergehause, and Geblase are not makers.
5. The first table column is often a drawing position even when its multilingual
   header contains "Part-No.". When it contains sequential callouts such as 1,
   2, 3, 1-12, P101, etc., return it as item_no.
5a. When one printed drawing position/item number is visually merged across two
   or more spare-part records, repeat that same item_no on EVERY returned record.
   Never leave a continuation record blank merely because the source cell is
   merged or printed only once.
6. ident_no is the identifier under headers such as Ident-Nr., Ident-No., Code,
   Material Code, Spare Part Code, or equivalent. This value will populate BOTH
   Benefit CODE and Benefit PART NO.
6a. Under a multilingual Bestell-Nr. / Order-No. / No de commande header,
    Order-No. is ident_no. Never confuse it with the Bild/Pict./Photo position,
    burner-series applicability, size, unit marker, or weight.
6b. OCR may flatten several Order-No. values into one cell. Return one distinct
    spare_parts record for every printed Order-No.; never combine them into one code.
7. source_part_no is any separate manufacturer Part No. printed by the source.
   Capture it only for audit context; it is not used automatically in Benefit.
8. description_english is the individual spare-part name only, in ENGLISH and
   UPPERCASE. When German/English/French are printed together, return ONLY the EXACT
   printed English phrase. Never concatenate the language variants. OCR may flatten
   the three printed lines into one string; still identify and isolate the English
   span. Do not paraphrase, modernize, shorten, expand, or replace it with a synonym.
   Preserve source dimensions, standards, stage numbers, punctuation, and symbols.
   When no English phrase is printed, translate the complete term into precise
   technical English and lower confidence because translation was required.
8a. Apply the same language rule to section_name_english: isolate the printed English
   title when present; otherwise translate the complete section title into English.
8b. Flattened OCR examples that MUST be cleaned:
   KURBELGEHÄUSE CRANK CASE CARTER -> CRANK CASE
   WELLENDICHTRING RADIAL PACKING RING BAGUE D'ETANCHEITE -> RADIAL PACKING RING
   LEISTUNGSSCHILD MANUFACTURER'S NAME PLATE PLAQUE SIGNALETIQUE -> MANUFACTURER'S NAME PLATE
   MANOMETER 1. STUFE PRESSURE GAUGE 1ST STAGE MANOMETRE 1ER ETAGE -> PRESSURE GAUGE 1ST STAGE
9. quantity must be numeric or null. unit is PCS, SET, or an empty string.
10. source_page is the PAGE marker supplied in the user message.
11. section_start_page is the first PDF page where the section begins, including
    its sectional/exploded drawing page when that page precedes the parts table.
12. Continuation pages keep exactly the same section_code, section_name_english,
    section_maker, section_model, and section_start_page only when the current page
    has no new drawing/table code in its own header. A code printed in a spare-part
    row, description, cross-reference, or table body is never a section_code.
12a. section_name_english contains only the clean English assembly title. Do not
     append section_code in parentheses and do not append ordering notes or the
     description of the first spare part below the heading.
13. Do not convert contents/index entries, drawing callouts without a parts table,
    page numbers, headers, or prose into spare-part rows.
14. confidence is 0 to 1 and reflects OCR quality, row alignment, English-language
    selection, and section matching.
15. A weight column headed ca. kg, appr. kg, env. kg, weight, or equivalent is not
    quantity. Return quantity=null when the source has no genuine quantity column.

Critical example:
Source columns:
  Teil-Nr./Part-No. | Ident-Nr./Ident-No. | Benennung/Designation | Qty
Source row:
  2 | 10.10.10.40 | Gleitlager / slide bearing / palier | 1
Return:
  item_no="2", ident_no="10.10.10.40",
  description_english="SLIDE BEARING", quantity=1.

Merged-position continuation example:
  2 | 53.20.1-.2 | 2/2-way solenoid valve 1st and 2nd stage | 2
    | 53.20.3    | 2/2-way solenoid valve 3rd stage         | 1
Return TWO rows and repeat item_no="2" on BOTH rows.
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


DOCUMENT_PROFILE_SYSTEM_PROMPT = """
You are a technical-document layout analyst. Examine representative OCR pages from
one marine or industrial spare-parts manual and describe the manual's extraction
pattern. Return exactly one JSON object with this structure:
{
  "document_family": "",
  "languages": [],
  "english_selection_rule": "",
  "part_identifier_header": "",
  "part_identifier_rule": "",
  "item_number_header": "",
  "item_number_rule": "",
  "quantity_header": "",
  "quantity_rule": "",
  "hierarchy": {
    "major_code_pattern": "",
    "subsection_code_pattern": "",
    "heading_detection_rule": "",
    "continuation_rule": "",
    "examples": [{"code": "", "name_english": "", "page": 1}]
  },
  "table_rules": [],
  "exclude_as_codes": [],
  "uncertainties": [],
  "confidence": 0.0
}

Rules:
1. Base every conclusion on text visibly present in the supplied OCR pages. Do not
   invent headings, translations, columns, codes, or page ranges.
2. Distinguish hierarchy codes from spare-part identifiers, drawing positions,
   burner/application variants, maker names, and model names such as SQM10 or ASZ12.
3. Preserve printed punctuation and trailing zeroes in hierarchy examples.
4. Explain how to select an already printed English phrase. Translation is allowed
   only when no English wording is printed.
5. Identify vertically merged cells, continuation pages, parallel language columns,
   repeated headers, weight columns, and multi-value Order-No./Ident-No. cells when
   the evidence supports them.
6. Keep rules concise and operational. Put any doubtful conclusion in uncertainties
   and lower confidence rather than guessing.
7. confidence is from 0 to 1 and describes confidence in the overall profile.
""".strip()


def _representative_profile_pages(
    extracted_pages: Sequence[tuple[int, str]],
    max_pages: int = 24,
    max_chars_per_page: int = 7000,
) -> list[tuple[int, str]]:
    """Select an evenly distributed, bounded sample for one document-level pass."""
    pages = [(int(page), str(markdown or "")) for page, markdown in extracted_pages]
    if len(pages) <= max(1, int(max_pages)):
        selected = pages
    else:
        last = len(pages) - 1
        indexes = {
            round(position * last / (max(2, int(max_pages)) - 1))
            for position in range(max(2, int(max_pages)))
        }
        selected = [pages[index] for index in sorted(indexes)]
    character_limit = max(1000, int(max_chars_per_page))
    return [
        (page, markdown[:character_limit])
        for page, markdown in selected
        if clean_text(markdown)
    ]


def _profile_text_list(value: Any, maximum: int = 20) -> list[str]:
    if not isinstance(value, list):
        return []
    result: list[str] = []
    for item in value:
        text = clean_text(item)
        if text and text not in result:
            result.append(text)
        if len(result) >= maximum:
            break
    return result


def _normalize_document_profile(
    value: Any,
    model: str,
    analyzed_pages: Sequence[int],
) -> dict[str, Any]:
    if not isinstance(value, dict):
        return {}
    hierarchy_source = value.get("hierarchy", {})
    hierarchy_source = hierarchy_source if isinstance(hierarchy_source, dict) else {}
    examples: list[dict[str, Any]] = []
    for item in hierarchy_source.get("examples", []):
        if not isinstance(item, dict):
            continue
        page_value = quantity_to_number(item.get("page"))
        examples.append(
            {
                "code": clean_text(item.get("code", "")),
                "name_english": clean_text(item.get("name_english", "")),
                "page": int(page_value) if page_value is not None else None,
            }
        )
        if len(examples) >= 30:
            break
    profile = {
        "document_family": clean_text(value.get("document_family", "")),
        "languages": _profile_text_list(value.get("languages", []), maximum=12),
        "english_selection_rule": clean_text(value.get("english_selection_rule", "")),
        "part_identifier_header": clean_text(value.get("part_identifier_header", "")),
        "part_identifier_rule": clean_text(value.get("part_identifier_rule", "")),
        "item_number_header": clean_text(value.get("item_number_header", "")),
        "item_number_rule": clean_text(value.get("item_number_rule", "")),
        "quantity_header": clean_text(value.get("quantity_header", "")),
        "quantity_rule": clean_text(value.get("quantity_rule", "")),
        "hierarchy": {
            "major_code_pattern": clean_text(hierarchy_source.get("major_code_pattern", "")),
            "subsection_code_pattern": clean_text(hierarchy_source.get("subsection_code_pattern", "")),
            "heading_detection_rule": clean_text(hierarchy_source.get("heading_detection_rule", "")),
            "continuation_rule": clean_text(hierarchy_source.get("continuation_rule", "")),
            "examples": examples,
        },
        "table_rules": _profile_text_list(value.get("table_rules", [])),
        "exclude_as_codes": _profile_text_list(value.get("exclude_as_codes", [])),
        "uncertainties": _profile_text_list(value.get("uncertainties", [])),
        "confidence": clamp_confidence(value.get("confidence"), fallback=0.0),
        "analysis_model": clean_text(model),
        "analyzed_pages": [int(page) for page in analyzed_pages],
    }
    meaningful = any(
        profile.get(key)
        for key in (
            "document_family", "languages", "part_identifier_header",
            "item_number_header", "table_rules",
        )
    ) or bool(examples)
    return profile if meaningful else {}


def analyze_document_profile_with_ai(
    api_key: str,
    model: str,
    extracted_pages: Sequence[tuple[int, str]],
    additional_instructions: str = "",
) -> tuple[dict[str, Any], list[str]]:
    """Create one evidence-based document profile before row extraction."""
    sample = _representative_profile_pages(extracted_pages)
    if not sample:
        raise ValueError("No OCR text was available for adaptive document analysis.")
    page_text = "\n\n".join(
        f"===== PDF PAGE {page} =====\n{markdown}"
        for page, markdown in sample
    )
    instruction_text = clean_text(additional_instructions)
    user_prompt = (
        "Analyze the following representative pages and return only the required "
        "document-profile JSON.\n\n"
    )
    if instruction_text:
        user_prompt += (
            "User-supplied manual notes (treat as guidance and verify against the OCR):\n"
            f"{instruction_text}\n\n"
        )
    user_prompt += page_text
    result = _mistral_json_request(
        api_key=api_key,
        model=clean_text(model) or "mistral-large-2512",
        messages=[
            {"role": "system", "content": DOCUMENT_PROFILE_SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        timeout_seconds=300,
        max_retries=3,
    )
    profile = _normalize_document_profile(
        result,
        clean_text(model) or "mistral-large-2512",
        [page for page, _ in sample],
    )
    if not profile:
        raise ValueError("Large-model analysis returned an empty document profile.")
    confidence = float(profile.get("confidence", 0.0))
    return profile, [
        "Adaptive document analysis completed with "
        f"{profile['analysis_model']} across {len(sample)} representative page(s) "
        f"(profile confidence {confidence:.0%})."
    ]


def _adaptive_profile_prompt(profile: dict[str, Any] | None) -> str:
    if not isinstance(profile, dict) or not profile:
        return ""
    prompt_profile = {
        key: profile.get(key)
        for key in (
            "document_family", "languages", "english_selection_rule",
            "part_identifier_header", "part_identifier_rule",
            "item_number_header", "item_number_rule", "quantity_header",
            "quantity_rule", "hierarchy", "table_rules", "exclude_as_codes",
            "uncertainties", "confidence",
        )
    }
    return (
        "ADAPTIVE DOCUMENT PROFILE (generated once from representative pages):\n"
        + json.dumps(prompt_profile, ensure_ascii=False, indent=2)
        + "\nUse this profile as document-wide guidance, but never override contradictory "
        "evidence on the current page or invent a value that is not printed."
    )


def _build_extraction_prompt(
    batch: Sequence[tuple[int, str]],
    additional_instructions: str,
    catalog_hint: str = "",
    profile_hint: str = "",
    adaptive_profile_hint: str = "",
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
    if profile_hint:
        prompt += profile_hint + "\n\n"
    if adaptive_profile_hint:
        prompt += adaptive_profile_hint + "\n\n"
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
    document_profile: dict[str, Any] | None = None,
    coverage_model: str = "",
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
    extraction_profile = _document_extraction_profile(extracted_pages)
    profile_hint = _profile_prompt(extraction_profile)
    adaptive_profile_hint = _adaptive_profile_prompt(document_profile)
    catalog_hint = _catalog_prompt_hint(
        build_section_catalog(extracted_pages).get("sections", [])
    )
    if extraction_profile == MULTILINGUAL_ORDER_CATALOG_PROFILE:
        messages.append(
            "Automatically detected a multilingual Order-No. catalogue. "
            "Order-No. is being mapped to PART NO/CODE, Pict./Photo to ITEM NO, "
            "and weight is being excluded from QNT."
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
                        "content": _build_extraction_prompt(
                            batch,
                            additional_instructions,
                            catalog_hint,
                            profile_hint,
                            adaptive_profile_hint,
                        ),
                    },
                ],
            )
            batch_rows = result.get("spare_parts", [])
            if not isinstance(batch_rows, list):
                raise ValueError("JSON did not contain a spare_parts list")
            normalized_batch_rows = _normalize_batch_source_pages(batch_rows, batch)

            # Some models return a valid but empty JSON object for dense historical
            # catalogues.  That is not an API failure, so the old recovery path never
            # ran. Retry once with an explicit coverage instruction before allowing
            # the deterministic Markdown parser to supplement the page later.
            expected_direct_rows = (
                _direct_table_rows(batch)
                if extraction_profile == MULTILINGUAL_ORDER_CATALOG_PROFILE
                else []
            )
            catalog_table_present = bool(
                extraction_profile == MULTILINGUAL_ORDER_CATALOG_PROFILE
                and any(
                    _is_multilingual_order_catalog(markdown)
                    for _, markdown in batch
                )
            )
            expected_count = len(expected_direct_rows)
            coverage_is_sparse = bool(
                catalog_table_present
                and (
                    len(normalized_batch_rows) < 2
                    or (
                        expected_count >= 3
                        and len(normalized_batch_rows) < max(2, int(expected_count * 0.85))
                    )
                )
            )
            if coverage_is_sparse:
                expected_wording = (
                    f"approximately {expected_count}"
                    if expected_direct_rows
                    else "multiple"
                )
                recovery_instructions = "\n\n".join(
                    value
                    for value in (
                        clean_text(additional_instructions),
                        (
                            "COVERAGE RECOVERY: The prior pass returned no rows even "
                            f"though the table contains {expected_wording} "
                            "Order-No. records. Read every table row and every Order-No. "
                            "variant. Do not summarize the page."
                        ),
                    )
                    if value
                )
                try:
                    retry = _mistral_json_request(
                        api_key=api_key,
                        model=model,
                        messages=[
                            {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                            {
                                "role": "user",
                                "content": _build_extraction_prompt(
                                    batch,
                                    recovery_instructions,
                                    catalog_hint,
                                    profile_hint,
                                    adaptive_profile_hint,
                                ),
                            },
                        ],
                    )
                    retry_rows = retry.get("spare_parts", [])
                    if isinstance(retry_rows, list):
                        normalized_retry_rows = _normalize_batch_source_pages(
                            retry_rows, batch
                        )
                        if len(normalized_retry_rows) > len(normalized_batch_rows):
                            normalized_batch_rows = normalized_retry_rows
                            messages.append(
                                f"Recovered {len(normalized_batch_rows)} row(s) from "
                                f"catalogue page(s) {batch[0][0]}-{batch[-1][0]} after "
                                "automatic coverage retry."
                            )
                except Exception as recovery_error:
                    messages.append(
                        f"Catalogue coverage retry was unavailable for page(s) "
                        f"{batch[0][0]}-{batch[-1][0]}; existing extracted rows were "
                        f"retained. Details: {recovery_error}"
                    )

            still_sparse = bool(
                catalog_table_present
                and (
                    len(normalized_batch_rows) < 2
                    or (
                        expected_count >= 3
                        and len(normalized_batch_rows) < max(2, int(expected_count * 0.85))
                    )
                )
            )
            recovery_model = clean_text(coverage_model)
            if (
                still_sparse
                and recovery_model
                and normalize_key(recovery_model) != normalize_key(model)
            ):
                large_recovery_instructions = "\n\n".join(
                    value
                    for value in (
                        clean_text(additional_instructions),
                        (
                            "HIGH-ACCURACY COVERAGE RECOVERY: Earlier extraction passes "
                            "missed records on this page. Reconstruct every visual table row, "
                            "split every Order-No. into its own spare-part record, preserve the "
                            "printed hierarchy heading, and do not promote a spare description "
                            "to section_name_english."
                        ),
                    )
                    if value
                )
                try:
                    large_retry = _mistral_json_request(
                        api_key=api_key,
                        model=recovery_model,
                        messages=[
                            {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                            {
                                "role": "user",
                                "content": _build_extraction_prompt(
                                    batch,
                                    large_recovery_instructions,
                                    catalog_hint,
                                    profile_hint,
                                    adaptive_profile_hint,
                                ),
                            },
                        ],
                    )
                    large_rows = large_retry.get("spare_parts", [])
                    if isinstance(large_rows, list):
                        normalized_large_rows = _normalize_batch_source_pages(
                            large_rows, batch
                        )
                        if len(normalized_large_rows) > len(normalized_batch_rows):
                            normalized_batch_rows = normalized_large_rows
                            messages.append(
                                f"Recovered {len(normalized_large_rows)} row(s) from "
                                f"catalogue page(s) {batch[0][0]}-{batch[-1][0]} with "
                                f"the high-accuracy model {recovery_model}."
                            )
                except Exception as recovery_error:
                    messages.append(
                        f"High-accuracy coverage recovery was unavailable for page(s) "
                        f"{batch[0][0]}-{batch[-1][0]}; existing extracted rows were "
                        f"retained. Details: {recovery_error}"
                    )

            rows.extend(normalized_batch_rows)
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




ENGLISH_NORMALIZATION_SYSTEM_PROMPT = """
You are a strict technical-language normalizer for marine and industrial spare-parts
manuals. Return one JSON object with exactly this structure:
{
  "items": [
    {"id": "D0", "english": ""}
  ]
}

For every input item:
1. Return only the English technical term in UPPERCASE.
2. The source may contain German, English, and French concatenated because OCR removed
   line breaks. Isolate the English wording; do not return all languages together.
3. When an English phrase is already printed, reproduce that phrase faithfully rather
   than replacing it with a synonym or paraphrase.
4. When no English wording is present, translate the complete term into precise
   technical English.
5. Remove German and French wording from the result. Keep proper names, maker/model
   names, dimensions, standards, symbols, stage numbers, and forms such as 2/2-WAY.
6. Never change identifiers, codes, item numbers, quantities, or dimensions.
7. Return exactly one result for every supplied id and no additional commentary.

Examples:
KURBELGEHÄUSE CRANK CASE CARTER -> CRANK CASE
WELLENDICHTRING RADIAL PACKING RING BAGUE D'ETANCHEITE -> RADIAL PACKING RING
LEISTUNGSSCHILD MANUFACTURER'S NAME PLATE PLAQUE SIGNALETIQUE -> MANUFACTURER'S NAME PLATE
MANOMETER 1. STUFE PRESSURE GAUGE 1ST STAGE MANOMETRE 1ER ETAGE -> PRESSURE GAUGE 1ST STAGE
ÖLSTANDSAUGE MIT DICHTRING -> OIL LEVEL SIGHT GLASS WITH SEALING RING
""".strip()


def _english_normalization_batches(
    items: Sequence[dict[str, Any]],
    max_items: int = 40,
    max_chars: int = 22000,
) -> list[list[dict[str, Any]]]:
    batches: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    current_chars = 0
    for item in items:
        item_chars = len(json.dumps(item, ensure_ascii=False))
        if current and (
            len(current) >= max(1, int(max_items))
            or current_chars + item_chars > max(2000, int(max_chars))
        ):
            batches.append(current)
            current = []
            current_chars = 0
        current.append(item)
        current_chars += item_chars
    if current:
        batches.append(current)
    return batches


def enforce_english_only_with_ai(
    api_key: str,
    model: str,
    rows: Sequence[dict[str, Any]],
    progress: ProgressCallback | None = None,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Isolate printed English or translate unresolved text before review/export.

    Descriptions are sent only when they are foreign, flattened multilingual text,
    or not confidently English. Every unique sub-machinery title is normalized once
    and then propagated by section code so one section cannot receive several names.
    """
    normalized_rows = [dict(row) for row in rows if isinstance(row, dict)]
    if not normalized_rows:
        return normalized_rows, []

    key = _normalize_api_key(api_key)
    request_items: list[dict[str, Any]] = []
    description_targets: dict[str, int] = {}
    section_targets: dict[str, str] = {}
    section_item_by_key: dict[str, str] = {}
    section_rows: dict[str, list[int]] = {}

    for row_index, row in enumerate(normalized_rows):
        current_description = clean_text(
            row.get("description_english", row.get("description", ""))
        )
        source_description = _clean_markdown_cell(
            row.get("source_description_raw", current_description)
        )
        if (
            clean_text(row.get("description_source", "")).lower() != "printed english"
            and (
                _text_needs_english_normalization(source_description)
                or _text_needs_english_normalization(current_description)
            )
        ):
            item_id = f"D{row_index}"
            request_items.append(
                {
                    "id": item_id,
                    "kind": "spare_part_description",
                    "source_text": source_description,
                    "current_text": current_description,
                    "section_code": clean_text(row.get("section_code", "")),
                    "item_no": clean_text(row.get("item_no", "")),
                    "ident_no": clean_text(row.get("ident_no", row.get("code", ""))),
                }
            )
            description_targets[item_id] = row_index

        section_code = clean_text(row.get("section_code", ""))
        current_section = clean_text(
            row.get("section_name_english", row.get("detected_machinery", ""))
        )
        source_section = _clean_markdown_cell(
            row.get("source_section_name_raw", row.get("table_title", current_section))
        )
        section_key = normalize_key(section_code) or f"NAME{normalize_key(current_section)}"
        section_rows.setdefault(section_key, []).append(row_index)
        section_needs_cleanup = (
            _text_needs_english_normalization(source_section)
            or _text_needs_english_normalization(current_section)
        )
        if (
            section_needs_cleanup
            and section_key not in section_item_by_key
            and (source_section or current_section)
        ):
            item_id = f"S{len(section_item_by_key)}"
            request_items.append(
                {
                    "id": item_id,
                    "kind": "sub_machinery_title",
                    "source_text": source_section or current_section,
                    "current_text": current_section,
                    "section_code": section_code,
                }
            )
            section_item_by_key[section_key] = item_id
            section_targets[item_id] = section_key

    if not request_items:
        return normalized_rows, [
            "All descriptions and sub-machinery titles were already confidently English."
        ]

    batches = _english_normalization_batches(request_items)
    returned: dict[str, str] = {}
    failed_ids: set[str] = set()

    def process_batch(batch: list[dict[str, Any]], label: str) -> None:
        try:
            result = _mistral_json_request(
                api_key=key,
                model=model or "mistral-small-latest",
                messages=[
                    {"role": "system", "content": ENGLISH_NORMALIZATION_SYSTEM_PROMPT},
                    {
                        "role": "user",
                        "content": (
                            "Normalize the following technical terms. Return only the required JSON object.\n\n"
                            + json.dumps({"items": batch}, ensure_ascii=False)
                        ),
                    },
                ],
                timeout_seconds=300,
                max_retries=3,
            )
            result_items = result.get("items", [])
            if not isinstance(result_items, list):
                raise ValueError("English normalization JSON did not contain an items list")
            batch_ids = {str(item.get("id", "")) for item in batch}
            seen_ids: set[str] = set()
            for result_item in result_items:
                if not isinstance(result_item, dict):
                    continue
                item_id = clean_text(result_item.get("id", ""))
                english = clean_text(result_item.get("english", "")).upper()
                if item_id in batch_ids and english:
                    returned[item_id] = english
                    seen_ids.add(item_id)
            missing = batch_ids - seen_ids
            if missing:
                raise ValueError(
                    "English normalization omitted item(s): " + ", ".join(sorted(missing))
                )
        except Exception:
            if len(batch) > 1:
                midpoint = max(1, len(batch) // 2)
                process_batch(batch[:midpoint], f"{label}.1")
                process_batch(batch[midpoint:], f"{label}.2")
            else:
                failed_ids.add(str(batch[0].get("id", "")))

    for batch_index, batch in enumerate(batches, start=1):
        if progress:
            progress(
                batch_index - 1,
                len(batches),
                f"English-only cleanup {batch_index}/{len(batches)}",
            )
        process_batch(batch, f"English batch {batch_index}")
        if progress:
            progress(
                batch_index,
                len(batches),
                f"English-only cleanup {batch_index}/{len(batches)} complete",
            )

    cleaned_descriptions = 0
    cleaned_sections = 0
    unresolved = 0

    for item_id, row_index in description_targets.items():
        english = returned.get(item_id, "")
        if not english:
            normalized_rows[row_index]["language_review"] = True
            unresolved += 1
            continue
        normalized_rows[row_index]["description_english"] = english
        normalized_rows[row_index]["description"] = english
        still_suspect = (
            _looks_likely_non_english(english)
            or len(_language_segments(english)) > 1
        )
        normalized_rows[row_index]["language_review"] = bool(still_suspect)
        normalized_rows[row_index]["english_cleanup_applied"] = True
        normalized_rows[row_index]["description_source"] = "AI English isolation/translation"
        if still_suspect:
            normalized_rows[row_index]["confidence"] = min(
                clamp_confidence(normalized_rows[row_index].get("confidence", 0.60)),
                0.55,
            )
            unresolved += 1
        else:
            # Translation improves language certainty but must not hide unrelated
            # section or OCR uncertainty.
            normalized_rows[row_index]["confidence"] = max(
                clamp_confidence(normalized_rows[row_index].get("confidence", 0.70)),
                0.72,
            )
            cleaned_descriptions += 1

    for item_id, section_key in section_targets.items():
        english = returned.get(item_id, "")
        if not english:
            unresolved += 1
            continue
        still_suspect = (
            _looks_likely_non_english(english)
            or len(_language_segments(english)) > 1
        )
        if still_suspect:
            unresolved += 1
            continue
        for row_index in section_rows.get(section_key, []):
            normalized_rows[row_index]["section_name_english"] = english
            normalized_rows[row_index]["detected_machinery"] = english
            normalized_rows[row_index]["section_english_cleanup_applied"] = True
        cleaned_sections += 1

    messages = [
        f"English-only cleanup normalized {cleaned_descriptions} spare-part description(s).",
        f"English-only cleanup normalized {cleaned_sections} unique sub-machinery title(s).",
    ]
    if unresolved or failed_ids:
        messages.append(
            f"{max(unresolved, len(failed_ids))} English cleanup item(s) remain flagged for review."
        )
    return normalized_rows, messages


# ---------------------------------------------------------------------------
# Deterministic Benefit table parsing and section catalogue
# ---------------------------------------------------------------------------


def _clean_markdown_cell(value: Any) -> str:
    """Clean a Markdown table cell while preserving meaningful line breaks.

    Mistral commonly separates German / English / French variants with ``<br>``.
    Keeping those boundaries lets the deterministic layer select the exact printed
    English phrase instead of trusting an AI paraphrase.
    """
    raw = "" if value is None else str(value)
    raw = raw.replace("\x00", " ").replace("\\|", "|")
    raw = re.sub(r"<br\s*/?>", "\n", raw, flags=re.IGNORECASE)
    raw = re.sub(r"[`*_]+", "", raw)
    lines = [re.sub(r"\s+", " ", line).strip() for line in raw.splitlines()]
    return "\n".join(line for line in lines if line)


def _split_markdown_row(line: str) -> list[str]:
    stripped = line.strip().strip("|")
    return [
        _clean_markdown_cell(cell)
        for cell in re.split(r"(?<!\\)\|", stripped)
    ]

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
    if any(
        token in key
        for token in ("orderno", "orderingno", "bestellnr", "nodecommande")
    ):
        return "ident_no"
    if "ident" in key or key in {"code", "partcode", "sparepartcode", "materialcode", "articlecode"}:
        return "ident_no"
    if any(token in key for token in ("description", "designation", "benennung", "partname", "denomination")):
        return "description_raw"
    if any(token in key for token in ("quantity", "qty", "qnt", "menge", "numberoff", "nooff")):
        return "quantity"
    if any(token in key for token in ("itemno", "itemnumber", "positionno", "position", "posno", "refno", "referenceno", "indexno")):
        return "item_no"
    if any(token in key for token in ("bildpictphoto", "pictphoto", "picturephoto")):
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
        has_identifier = any(
            token in joined
            for token in (
                "ident", "code", "part-no", "part no", "item", "position",
                "order-no", "order no", "bestell-nr", "bestell nr",
                "no de commande", "pict.", "photo",
            )
        )
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


def _language_segments(value: Any) -> list[str]:
    """Return visually separated language variants without splitting values like 2/2-way."""
    raw = _clean_markdown_cell(value)
    if not raw:
        return []
    line_parts = [line.strip() for line in raw.splitlines() if line.strip()]
    if len(line_parts) >= 2:
        return line_parts
    # Only split slashes surrounded by whitespace. Technical values such as
    # ``2/2-way`` and dimensions such as ``10/12 mm`` must remain intact.
    slash_parts = [
        clean_text(part)
        for part in re.split(r"\s+/\s+", raw)
        if clean_text(part)
    ]
    return slash_parts if len(slash_parts) >= 2 else [clean_text(raw)]


def _looks_likely_non_english(value: Any) -> bool:
    text = clean_text(value).lower()
    markers = (
        # German
        "schraube", "mutter", "dichtung", "gleitlager", "stufe", "zylinder",
        "halter", "kupplung", "scheibe", "leitung", "gehäuse", "gehaeuse",
        "kurbelgehäuse", "kurbelgehaeuse", "kurbelwelle", "wellendichtring",
        "leistungs", "drehrichtung", "einschraub", "ölstand", "oelstand",
        "dichtring", "passfeder", "spannhülse", "spannhuelse", "deckel",
        "lager", "welle", "stutzen", "druckluft", "sicherungsring",
        "sechskant", "manometer", "benennung", "abdeckung", "drossel",
        # French
        "pièce", "piece de", "soupape", "tuyau", "arbre", "carter", "écrou",
        "ecrou", "palier", "étage", "etage", "raccord", "support pour",
        "vis hexagonale", "joint de", "bague", "couvercle", "plaque",
        "flèche", "fleche", "rondelle", "clavette", "manchon", "vilebrequin",
        "étanch", "etanch", "graisseur", "ressort", "tôle", "tole",
        "désignation", "designation française", "quantité", "quantite",
        "brûleur", "bruleur", "moteur", " pour ", "réglage", "reglage",
        "brenner", "gebläse", "geblase", "stellantrieb", "luftregelung",
        "pumpen", "druckwächter", "druckwachter", "flammkopf", "düsen",
        "dusen", "zündtrafo", "zundtrafo", "ölvorwärmer", "olvorwarmer",
    )
    return any(marker in text for marker in markers)


def _looks_likely_english(value: Any) -> bool:
    text = clean_text(value).lower()
    if not text or _looks_likely_non_english(text):
        return False
    markers = (
        " for ", " with ", " and ", " of ", "stage", "support", "motor",
        "bolt", "nut", "ring", "valve", "pipe", "gauge", "piece", "hose",
        "bearing", "coupling", "mounting", "separator", "filter", "plate",
        "disc", "cylinder", "head", "pressure", "suction", "flywheel",
        "socket", "clip", "flange", "fitting", "thermometer", "gasket",
        "joint", "seal", "spring", "shaft", "cover", "piston", "liner",
        "screw", "washer", "nozzle", "bracket", "cartridge", "crank",
        "case", "radial", "packing", "inspection", "hole", "manufacturer",
        "name", "arrow", "direction", "oil", "level", "sight", "glass",
        "straight", "male", "union", "gear", "wheel", "complete", "dowel",
        "key", "cap", "tube", "sieve", "seat", "guide", "rod", "lever",
        "handle", "body", "housing", "insert", "bush", "bushing", "pin",
        "plug", "stud", "panel", "rubber", "metal", "foot", "inside",
        "outside", "adjustable", "protecting", "line", "air", "compressed",
        "elbow", "reduction", "reducing", "threaded", "filter", "cartridge",
        "burner", "casing", "individual", "regulation", "servo", "drive",
        "blower", "pump", "combustion", "ignition", "transformer", "preheater",
        "monitor", "terminal", "electrical", "equipment", "train",
    )
    padded = f" {text} "
    return any(marker in padded for marker in markers)


def _text_needs_english_normalization(value: Any) -> bool:
    """Return True when text is foreign, multilingual, or not confidently English."""
    raw = _clean_markdown_cell(value)
    text = clean_text(raw)
    if not text:
        return True
    if len(_language_segments(raw)) >= 2:
        return True
    if _looks_likely_non_english(text):
        return True
    # Accented Latin characters are a strong signal in German/French/English manuals.
    if re.search(r"[À-ÿ]", text):
        return True
    return not _looks_likely_english(text)

def _english_variant(value: Any) -> str:
    segments = _language_segments(value)
    if not segments:
        return ""
    if len(segments) >= 3:
        return clean_text(segments[1]).upper()
    if len(segments) == 2:
        first, second = segments
        if _looks_likely_non_english(first) and _looks_likely_english(second):
            return clean_text(second).upper()
    return clean_text(segments[0]).upper()


_CATALOG_HEADING_CODE_RE = re.compile(
    r"^\s*(\d{1,2}(?:\.\d{1,2})?)\.?\s*$",
    flags=re.IGNORECASE,
)
_SIMPLE_SECTION_HEADING_RE = re.compile(
    r"^\s*(\d{1,2}(?:\.\d{1,2})?)\.?\s+(.+?[A-Za-zÀ-ÿ].*)$",
    flags=re.IGNORECASE,
)


def _catalog_heading_title(value: Any) -> str:
    """Return only the leading heading from a vertically merged title cell.

    These catalogues frequently merge a section title with the first spare-part
    description.  A wrapped title is joined only when the grammar clearly
    continues (``Burner casing and`` + ``individual parts``); later notes such as
    ``When ordering...`` or child descriptions such as ``Blower wheel`` are not.
    """
    lines = [
        re.sub(_SIMPLE_SECTION_HEADING_RE, r"\2", line).strip(" -:;|")
        for line in _catalog_cell_lines(value)
    ]
    lines = [line for line in lines if line and re.search(r"[A-Za-zÀ-ÿ]", line)]
    if not lines:
        return ""

    title_parts = [lines[0]]
    for line in lines[1:]:
        current = title_parts[-1]
        repeated = normalize_key(line).startswith(normalize_key(title_parts[0]))
        continues = bool(
            re.search(r"(?:\b(?:and|or|for|of|with|without|to)|[-/])$", current, re.I)
            or re.match(r"^(?:and|or|for|of|with|without|to)\b", line, re.I)
        )
        if repeated or not continues:
            break
        title_parts.append(line)
    return clean_text(" ".join(title_parts)).strip(" -:;|")


def _clean_catalog_section_name(value: Any, section_code: Any = "") -> str:
    """Remove OCR spillover and code labels from an assembly heading."""
    text = clean_text(value).strip(" -:;|").upper()
    code = clean_text(section_code).strip().rstrip(".")
    if not text:
        return ""
    if code:
        text = re.sub(
            rf"\s*\(\s*{re.escape(code)}\s*\)\s*$", "", text, flags=re.I
        ).strip()
        text = re.sub(
            rf"^\s*{re.escape(code)}\.?\s+", "", text, flags=re.I
        ).strip()

    # Ordering notes and the first child description are commonly flattened onto
    # the heading. They are not part of the sub-machinery name.
    text = re.split(r"\bWHEN\s+ORDERING\b", text, maxsplit=1, flags=re.I)[0].strip()
    words = text.split()
    for length in range(min(5, len(words) // 2), 0, -1):
        if words[:length] == words[length : length * 2]:
            text = " ".join(words[:length])
            break
    return clean_text(text).strip(" -:;|").upper()


def _clean_catalog_maker(value: Any) -> str:
    """Reject German assembly headings that OCR/AI can mislabel as makers."""
    text = clean_text(value).upper()
    compact = normalize_key(text)
    false_makers = {
        "BRENNERMOTOR",
        "BRENNERGEHAUSE",
        "BRENNERGEHAEUSE",
        "LUFTREGELUNG",
        "STELLANTRIEB",
        "GEBLASE",
        "GEBLAESE",
    }
    if compact in false_makers:
        return ""
    return text


def _catalog_row_without_heading(
    row: Sequence[Any], mappings: Sequence[str | None], section: dict[str, str]
) -> list[Any]:
    """Remove a merged heading prefix while retaining the first spare row."""
    values = list(row)
    code_key = normalize_key(section.get("section_code", ""))
    for index, canonical in enumerate(mappings):
        if index >= len(values):
            continue
        if canonical in {"item_no", "source_part_no"}:
            lines = _catalog_cell_lines(values[index])
            if lines:
                match = _CATALOG_HEADING_CODE_RE.fullmatch(lines[0])
                if match and normalize_key(match.group(1)) == code_key:
                    values[index] = "\n".join(lines[1:])
        elif canonical == "description_raw":
            lines = _catalog_cell_lines(values[index])
            title = _catalog_heading_title(values[index])
            if not lines or not title:
                continue
            consumed = 0
            rebuilt = ""
            for consumed_count in range(1, min(3, len(lines)) + 1):
                rebuilt = clean_text(" ".join(lines[:consumed_count]))
                if normalize_key(rebuilt) == normalize_key(title):
                    consumed = consumed_count
                    break
            if consumed:
                values[index] = "\n".join(lines[consumed:])
    return values


def _catalog_heading_code(value: Any) -> tuple[str, bool]:
    """Return a printed hierarchy code and whether the cell is clearly a heading."""
    lines = _catalog_cell_lines(value)
    if not lines:
        return "", False
    embedded = _SIMPLE_SECTION_HEADING_RE.match(lines[0])
    if embedded:
        return embedded.group(1), True

    codes: list[str] = []
    for line in lines:
        match = _CATALOG_HEADING_CODE_RE.fullmatch(line)
        if match:
            codes.append(match.group(1))
    if not codes:
        return "", False
    # ``1.`` is an unambiguous major heading. Repeated positions such as
    # ``12.1 / 12.1 / 12.1`` are applicability variants of one spare, not a
    # hierarchy. Decimal subsection headings in this catalogue use decade
    # boundaries (3.30, 6.40, 18.50) followed by distinct child positions.
    major_heading = bool(re.fullmatch(r"\s*\d{1,2}\.\s*", lines[0]))
    first_parts = codes[0].split(".")
    distinct_codes = list(dict.fromkeys(codes))
    decimal_heading = False
    if len(first_parts) == 2 and first_parts[1].isdigit():
        major, position = int(first_parts[0]), int(first_parts[1])
        decimal_heading = bool(
            position >= 10
            and position % 10 == 0
            and any(
                len(parts := candidate.split(".")) == 2
                and parts[0].isdigit()
                and parts[1].isdigit()
                and int(parts[0]) == major
                and int(parts[1]) > position
                for candidate in distinct_codes[1:]
            )
        )
    return codes[0], bool(major_heading or decimal_heading)


def _classic_catalog_section_from_row(
    row: Sequence[Any], mappings: Sequence[str | None] | None = None
) -> dict[str, str]:
    """Extract one major/subsection heading from a multilingual catalogue row."""
    values = list(row)
    code = ""
    strong_heading = False
    english = ""

    if mappings:
        item_indexes = [
            index
            for index, canonical in enumerate(mappings)
            if canonical in {"item_no", "source_part_no"}
        ]
        for index in item_indexes:
            if index >= len(values):
                continue
            candidate_code, candidate_is_heading = _catalog_heading_code(values[index])
            if candidate_code:
                code = candidate_code
                strong_heading = candidate_is_heading
                break

        ident_indexes = [
            index for index, canonical in enumerate(mappings) if canonical == "ident_no"
        ]
        has_order_number_text = any(
            index < len(values) and bool(clean_text(values[index]))
            for index in ident_indexes
        )
        # A decimal drawing position with an Order-No. is a spare row, not a
        # subsection. Merged heading/child rows remain valid when their item cell
        # proves that a heading precedes one or more child positions.
        if not code or (has_order_number_text and not strong_heading):
            return {}

        english_index = _classic_english_description_index(values, mappings)
        if english_index is not None and english_index < len(values):
            english = _catalog_heading_title(values[english_index])

    raw_candidates: list[str] = []
    for cell in values:
        segments = _language_segments(cell)
        for segment in segments:
            match = _SIMPLE_SECTION_HEADING_RE.match(segment)
            if match:
                if not code:
                    code = match.group(1)
                title = clean_text(match.group(2))
                if title:
                    raw_candidates.append(title)
            elif not mappings and re.fullmatch(r"\s*\d{1,2}\.\s*", segment):
                code = re.sub(r"\D", "", segment)
            elif code and re.search(r"[A-Za-zÀ-ÿ]", segment):
                raw_candidates.append(clean_text(segment))

    if not code:
        joined = " ".join(clean_text(cell) for cell in values if clean_text(cell))
        match = _SIMPLE_SECTION_HEADING_RE.match(joined)
        if match:
            code = match.group(1)
            raw_candidates.append(clean_text(match.group(2)))

    candidates: list[str] = []
    for value in raw_candidates:
        cleaned = clean_text(value).strip(" -:;|")
        lowered = cleaned.lower()
        if not cleaned or any(
            token in lowered
            for token in (
                "order-no", "bestell", "no de commande", "designation",
                "bezeichnung", "burner serie", "type bruleur", "type brûleur",
                "appr. kg", "env. kg",
            )
        ):
            continue
        if cleaned not in candidates:
            candidates.append(cleaned)

    if not english:
        english_candidates = [
            value
            for value in candidates
            if _looks_likely_english(value) and not _looks_likely_non_english(value)
        ]
        if english_candidates:
            english = english_candidates[0]
        elif len(candidates) >= 3:
            english = candidates[1]
        elif candidates:
            english = _english_variant("\n".join(candidates)) or candidates[0]

    if not code or not english:
        return {}
    return {
        "section_code": code,
        "section_name_raw": "\n".join(candidates) or english,
        "section_name_english": _clean_catalog_section_name(english, code),
    }


def _classic_catalog_sections(markdown: str) -> list[dict[str, str]]:
    """Return every numbered section heading, including two sections on one page."""
    if not _is_multilingual_order_catalog(markdown):
        return []
    result: list[dict[str, str]] = []
    seen: set[str] = set()

    # OCR sometimes places the real section heading immediately above the
    # Markdown table instead of inside its first row. Read those standalone
    # headings even when the same page also contains another valid table heading.
    for line in str(markdown or "").splitlines()[:100]:
        if "|" in line or line.lstrip().startswith("!["):
            continue
        cleaned = _clean_markdown_cell(line.lstrip("# "))
        if not cleaned:
            continue
        section = _classic_catalog_section_from_row([cleaned])
        code = section.get("section_code", "")
        if code and code not in seen:
            seen.add(code)
            result.append(section)

    for block in _markdown_table_blocks(markdown):
        header_index = _find_data_header_index(block)
        if header_index is None:
            continue
        headers = block[header_index]
        mappings = [_canonical_source_header(header) for header in headers]
        for row in block[header_index + 1 :]:
            section = _classic_catalog_section_from_row(row, mappings)
            code = section.get("section_code", "")
            if code and code not in seen:
                seen.add(code)
                result.append(section)
    return result


def _classic_catalog_section_metadata(markdown: str) -> dict[str, str]:
    """Return the first numbered section heading for page-level continuation."""
    sections = _classic_catalog_sections(markdown)
    return sections[0] if sections else {}


def _best_effort_english_description(value: Any) -> tuple[str, bool]:
    """Return uppercase English text and whether language review is still needed."""
    segments = _language_segments(value)
    if not segments:
        return "", True
    if len(segments) >= 3:
        return clean_text(segments[1]).upper(), False
    if len(segments) == 2:
        first, second = segments
        if _looks_likely_non_english(first) and _looks_likely_english(second):
            return clean_text(second).upper(), False
    text = clean_text(segments[0])
    # A single AI phrase is accepted only when it looks plausibly English. This
    # catches cases such as MANOMETER / STUFE being returned as the English line.
    needs_review = _looks_likely_non_english(text) or not _looks_likely_english(text)
    return text.upper(), needs_review


def _source_english_description(value: Any) -> tuple[str, bool]:
    """Return an exact English phrase printed in the source when it can be isolated."""
    segments = _language_segments(value)
    if not segments:
        return "", False
    if len(segments) >= 3:
        return clean_text(segments[1]).upper(), True
    if len(segments) == 2:
        first, second = segments
        if _looks_likely_non_english(first) and _looks_likely_english(second):
            return clean_text(second).upper(), True
    if len(segments) == 1 and _looks_likely_english(segments[0]):
        return clean_text(segments[0]).upper(), True
    return "", False


def _select_source_faithful_description(
    source_value: Any,
    ai_value: Any,
) -> tuple[str, bool, bool, str]:
    """Choose a description while keeping exact source English ahead of AI wording.

    Returns ``(description, needs_review, ai_disagreed, source_kind)``.
    """
    source_description, source_is_exact = _source_english_description(source_value)
    ai_description, ai_needs_review = _best_effort_english_description(ai_value)

    if source_is_exact and source_description:
        disagreed = bool(
            ai_description
            and SequenceMatcher(
                None,
                normalize_key(source_description),
                normalize_key(ai_description),
            ).ratio() < 0.92
        )
        # The printed English phrase is authoritative even when the AI produced a
        # fluent synonym or paraphrase.
        return source_description, False, disagreed, "printed English"

    if ai_description:
        return ai_description, bool(ai_needs_review), False, "AI English/translation"

    fallback = clean_text(source_value).upper()
    return fallback, True, False, "unresolved source text"

def _page_header_text(markdown: str, max_lines: int = 28) -> str:
    """Return only the page header / pre-table area used for section detection."""
    candidates: list[str] = []
    for block in _markdown_table_blocks(markdown):
        header_index = _find_data_header_index(block)
        if header_index is not None and header_index > 0:
            for row in block[:header_index]:
                candidates.extend(cell for cell in row if clean_text(cell))
            break

    if not candidates:
        for line in str(markdown or "").splitlines():
            cleaned = _clean_markdown_cell(line.lstrip("# "))
            if not cleaned or line.lstrip().startswith("!["):
                continue
            lowered = clean_text(cleaned).lower()
            if (
                any(token in lowered for token in ("description", "designation", "benennung"))
                and any(
                    token in lowered
                    for token in (
                        "ident", "part-no", "part no", "item", "order-no",
                        "order no", "bestell", "no de commande", "pict", "photo",
                    )
                )
            ):
                break
            candidates.append(cleaned)
            if len(candidates) >= max_lines:
                break
    return "\n".join(dict.fromkeys(clean_text(value) for value in candidates if clean_text(value)))


def _page_metadata(page_number: int, markdown: str) -> dict[str, Any]:
    metadata: dict[str, Any] = {
        "page": int(page_number),
        "section_code": "",
        "section_name_raw": "",
        "section_name_english": "",
        "maker": "",
        "model": "",
        "header_text": "",
        "section_code_source": "",
    }
    classic_catalog = _is_multilingual_order_catalog(markdown)
    blocks = _markdown_table_blocks(markdown)
    for block in blocks:
        header_index = _find_data_header_index(block)
        if header_index is None or header_index <= 0:
            continue
        preheader_rows = block[:header_index]
        meta_row = max(
            preheader_rows,
            key=lambda row: sum(bool(clean_text(value)) for value in row),
        )
        nonempty = [value for value in meta_row if clean_text(value)]
        if nonempty:
            metadata["maker"] = clean_text(nonempty[0]).upper()

        tail = nonempty[-1] if nonempty else ""
        tail_tokens = _section_code_tokens(tail, permissive=True)
        if tail_tokens:
            metadata["section_code"] = _join_section_codes(tail_tokens)
            metadata["section_code_source"] = "page header cell"

        middle_candidates = nonempty[1:-1] if len(nonempty) >= 3 else nonempty[1:]
        title_candidates = []
        for value in middle_candidates:
            cleaned = _clean_markdown_cell(value)
            flat = clean_text(cleaned)
            if not re.search(r"[A-Za-zÀ-ÿ]", flat):
                continue
            if re.fullmatch(r"[A-Z]?\s*\d+(?:[ ._/-]\d+)*", flat, flags=re.I):
                continue
            if re.search(r"\b(?:TAFEL|TABLE|PLANCHE|DRAWING|DWG)\b", flat, flags=re.I):
                continue
            title_candidates.append(cleaned)
        if title_candidates:
            # Prefer the richest multilingual title cell rather than a short model label.
            metadata["section_name_raw"] = max(
                title_candidates,
                key=lambda value: (len(_language_segments(value)), len(clean_text(value))),
            )

        model_part = re.split(
            r"\b(?:TAFEL|TABLE|PLANCHE|DRAWING|DWG)\b",
            clean_text(tail),
            maxsplit=1,
            flags=re.I,
        )[0]
        metadata["model"] = clean_text(model_part).upper()
        break

    header_text = _page_header_text(markdown)
    metadata["header_text"] = header_text
    if not metadata["section_code"] and not classic_catalog:
        # Never search the spare-parts table body for a section code. Identifiers and
        # cross-references in the body caused the previous section to remain active.
        tokens = _section_code_tokens(header_text, permissive=True)
        if tokens:
            metadata["section_code"] = _join_section_codes(tokens)
            metadata["section_code_source"] = "page header area"

    text = str(markdown or "")
    if not metadata["maker"]:
        bold_candidates = re.findall(r"\*\*([^*]{2,60})\*\*", header_text or text)
        for candidate in bold_candidates:
            candidate_clean = clean_text(candidate).upper()
            if re.search(r"[A-Z]", candidate_clean) and not re.search(r"\d", candidate_clean):
                if candidate_clean not in {"DESCRIPTION", "DESIGNATION", "QUANTITY", "TABLE"}:
                    metadata["maker"] = candidate_clean
                    break

    if not metadata["section_name_raw"]:
        lines = [
            _clean_markdown_cell(line.lstrip("# "))
            for line in (header_text or text).splitlines()
            if clean_text(line.lstrip("# ")) and not line.lstrip().startswith("![")
        ]
        stop_index = next(
            (idx for idx, line in enumerate(lines) if re.search(r"\b(?:TAFEL|TABLE|PLANCHE)\b", clean_text(line), flags=re.I)),
            None,
        )
        title_lines = lines[:stop_index] if stop_index is not None else lines[:6]
        title_lines = [line for line in title_lines if not _section_code_tokens(line)]
        title_lines = [
            line for line in title_lines
            if not re.fullmatch(r"[A-Z]?\s*\d+(?:\s*-\s*[A-Z]?\s*\d+)*", clean_text(line), flags=re.I)
        ]
        if len(title_lines) >= 3:
            metadata["section_name_raw"] = "\n".join(title_lines[:3])
        elif title_lines:
            metadata["section_name_raw"] = title_lines[0]

    metadata["section_name_english"] = _english_variant(metadata["section_name_raw"])
    if classic_catalog:
        classic_section = _classic_catalog_section_metadata(markdown)
        # In this layout, only a simple numbered heading is authoritative. Clear
        # false codes such as PG11/ASZ12 or an Order-No. read from the table body.
        metadata["section_code"] = classic_section.get("section_code", "")
        metadata["section_code_source"] = (
            "numbered catalogue section heading" if classic_section else ""
        )
        if classic_section:
            metadata["section_name_raw"] = classic_section["section_name_raw"]
            metadata["section_name_english"] = classic_section[
                "section_name_english"
            ]
        else:
            metadata["section_name_raw"] = ""
            metadata["section_name_english"] = ""
        # Words such as Brennermotor, Luftregelung and Stellantrieb are German
        # assembly titles in this layout, not manufacturer/model metadata. The
        # verified main-machinery maker/model are applied later as the fallback.
        metadata["maker"] = ""
        metadata["model"] = ""
    if not metadata["model"] and not classic_catalog:
        model_match = re.search(
            r"\b(?:COMPRESSOR|KOMPRESSOR|ENGINE|GENERATOR|PUMP)\s+([A-Z0-9][A-Z0-9 ._/-]{1,30})",
            header_text or text,
            flags=re.I,
        )
        if model_match:
            metadata["model"] = clean_text(model_match.group(1)).upper()
    return metadata

def _catalog_cell_lines(value: Any) -> list[str]:
    raw = _clean_markdown_cell(value)
    return [clean_text(line) for line in raw.splitlines() if clean_text(line)]


def _catalog_order_number_lines(value: Any) -> list[str]:
    """Return individual Order-No. values from a vertically merged catalogue cell."""
    result: list[str] = []
    for line in _catalog_cell_lines(value):
        candidate = re.sub(r"\s+[A-Z]$", "", line.strip(), flags=re.IGNORECASE)
        candidate = re.sub(r"\s+", " ", candidate).strip(" ;,|")
        # OCR can flatten several visual Order-No. lines into one text line. Split
        # the known manufacturer formats before applying the generic fallback.
        tokens = re.findall(
            r"(?<![\d/])(?:"
            r"\d{3,4}\s+\d{3}\s+\d{3,4}/\d{1,2}"
            r"|\d{9,12}/\d{1,2}"
            r"|\d{3,4}[ .]\d{3}"
            r"|\d{6}"
            r")(?![\d/])",
            candidate,
        )
        if tokens:
            for token in tokens:
                cleaned_token = re.sub(r"\s+", " ", token).strip(" .;,|")
                if cleaned_token and cleaned_token not in result:
                    result.append(cleaned_token)
            continue
        digit_count = len(re.findall(r"\d", candidate))
        if digit_count < 5:
            continue
        # Reject weights and simple sizes while accepting values such as
        # 499 089, 151 327 1512/2, and 110 500 0009/2.
        if re.fullmatch(r"\d+[,.]\d+", candidate):
            continue
        if not (re.search(r"\s", candidate) or "/" in candidate or "-" in candidate):
            continue
        if candidate not in result:
            result.append(candidate)
    return result


def _classic_english_description_index(
    headers: Sequence[str], mappings: Sequence[str | None]
) -> int | None:
    description_indexes = [
        index for index, canonical in enumerate(mappings) if canonical == "description_raw"
    ]
    if not description_indexes:
        return None
    # German / English / French are printed from left to right. Bezeichnung is
    # normally the first description column and the English Designation is second.
    if len(description_indexes) >= 3:
        return description_indexes[1]
    if len(description_indexes) == 2:
        # The German Bezeichnung often shares the applicability column and is
        # therefore not classified as a description column. The two remaining
        # Designation columns are English first, French second.
        return description_indexes[0]
    return description_indexes[0]


def _direct_table_rows(extracted_pages: Sequence[tuple[int, str]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for page_number, markdown in extracted_pages:
        classic_catalog = _is_multilingual_order_catalog(markdown)
        metadata = _page_metadata(int(page_number), markdown)
        for block in _markdown_table_blocks(markdown):
            header_index = _find_data_header_index(block)
            if header_index is None:
                continue
            headers = block[header_index]
            mappings = [_canonical_source_header(header) for header in headers]
            if classic_catalog:
                english_index = _classic_english_description_index(headers, mappings)
                mappings = [
                    (
                        canonical
                        if canonical != "description_raw" or index == english_index
                        else None
                    )
                    for index, canonical in enumerate(mappings)
                ]
            # When Ident-No. exists, a separate Part-No. column containing drawing
            # callouts is ITEM NO by the user's Benefit mapping rule.
            if "ident_no" in mappings and "item_no" not in mappings and "source_part_no" in mappings:
                mappings[mappings.index("source_part_no")] = "item_no"
            active_metadata = dict(metadata)
            for values in block[header_index + 1 :]:
                if classic_catalog:
                    row_section = _classic_catalog_section_from_row(values, mappings)
                    if row_section:
                        active_metadata.update(row_section)
                        ident_indexes = [
                            index
                            for index, canonical in enumerate(mappings)
                            if canonical == "ident_no"
                        ]
                        if not any(
                            index < len(values) and bool(clean_text(values[index]))
                            for index in ident_indexes
                        ):
                            continue
                        values = _catalog_row_without_heading(
                            values, mappings, row_section
                        )
                padded = list(values) + [""] * max(0, len(headers) - len(values))
                record: dict[str, Any] = {
                    "source_page": int(page_number),
                    "section_code": active_metadata.get("section_code", ""),
                    "section_name_english": active_metadata.get("section_name_english", ""),
                    "section_maker": active_metadata.get("maker", ""),
                    "section_model": active_metadata.get("model", ""),
                    "table_title": active_metadata.get("section_name_raw", ""),
                    "confidence": 0.88,
                }
                for column_index, canonical in enumerate(mappings):
                    if canonical and column_index < len(padded):
                        value = padded[column_index]
                        record[canonical] = (
                            _clean_markdown_cell(value)
                            if canonical in {"description_raw", "ident_no"}
                            else clean_text(value)
                        )

                if classic_catalog:
                    description_raw = _clean_markdown_cell(
                        record.get("description_raw", "")
                    )
                    item_no = clean_text(record.get("item_no", ""))
                    identifiers = _catalog_order_number_lines(
                        record.get("ident_no", "")
                    )
                    if not description_raw or not identifiers:
                        continue
                    for identifier in identifiers:
                        variant = dict(record)
                        variant["ident_no"] = identifier
                        variant["item_no"] = item_no
                        variant["quantity"] = None
                        variant["confidence"] = 0.92
                        rows.append(variant)
                    continue

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
    has_identifier = any(
        token in text
        for token in (
            "ident-nr", "ident-no", "ident no", "item no", "position",
            "part no", "part-no", "code", "order-no", "order no",
            "bestell-nr", "bestell nr", "no de commande", "pict.",
        )
    )
    has_quantity = any(token in text for token in ("qty", "quantity", "menge", "qnt"))
    return has_description and has_identifier and (
        has_quantity
        or "|" in str(markdown or "")
        or _is_multilingual_order_catalog(markdown)
    )


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
    """Build a section catalogue anchored to codes printed in each page header.

    Version 4.7 continues to avoid scanning the spare-parts table body for section
    codes. Body identifiers and cross-references previously produced multiple matches,
    leaving the previous ``active`` section stuck while the spare rows moved on.
    """
    extraction_profile = _document_extraction_profile(extracted_pages)
    allow_simple_section_codes = (
        extraction_profile == MULTILINGUAL_ORDER_CATALOG_PROFILE
    )
    index_sections, index_pages = _index_sections(extracted_pages)
    global_maker, global_model = _global_manual_maker_model(extracted_pages)
    main_row = main_row or {}
    sections: list[dict[str, Any]] = []
    by_alias: dict[str, dict[str, Any]] = {}

    def catalog_code_tokens(value: Any) -> list[str]:
        text = clean_text(value).upper().strip().rstrip(".")
        if allow_simple_section_codes:
            # This profile prints numeric hierarchy codes. Model names such as
            # SQM10 and ASZ12 are applicability data, never sub-machinery codes.
            return (
                [text]
                if re.fullmatch(r"\d{1,2}(?:\.\d{1,2})?", text)
                else []
            )
        return _section_code_tokens(value, permissive=True)

    def register(section: dict[str, Any], priority: int = 50) -> dict[str, Any]:
        aliases = [
            clean_text(value).upper()
            for value in section.get("aliases", [])
            if clean_text(value)
        ]
        code = clean_text(section.get("code", "")).upper() or _join_section_codes(aliases)
        if not aliases and code:
            aliases = catalog_code_tokens(code) or [code]
        aliases = list(dict.fromkeys(aliases))
        name = _clean_catalog_section_name(section.get("name", ""), code)
        existing = next(
            (
                by_alias.get(normalize_key(alias))
                for alias in aliases
                if by_alias.get(normalize_key(alias)) is not None
            ),
            None,
        )
        if existing is None:
            existing = {
                "code": code,
                "aliases": aliases or ([code] if code else []),
                "name": name,
                "maker": _clean_catalog_maker(section.get("maker", ""))
                if allow_simple_section_codes
                else clean_text(section.get("maker", "")).upper(),
                "model": clean_text(section.get("model", "")).upper(),
                "pages": set(section.get("pages", set())),
                "_name_priority": int(priority if name else -1),
                "_maker_priority": int(
                    priority
                    if (
                        _clean_catalog_maker(section.get("maker", ""))
                        if allow_simple_section_codes
                        else clean_text(section.get("maker", ""))
                    )
                    else -1
                ),
                "_model_priority": int(priority if clean_text(section.get("model", "")) else -1),
            }
            sections.append(existing)
        else:
            if name and int(priority) > int(existing.get("_name_priority", -1)):
                existing["name"] = name
                existing["_name_priority"] = int(priority)
            maker = (
                _clean_catalog_maker(section.get("maker", ""))
                if allow_simple_section_codes
                else clean_text(section.get("maker", "")).upper()
            )
            if maker and int(priority) > int(existing.get("_maker_priority", -1)):
                existing["maker"] = maker
                existing["_maker_priority"] = int(priority)
            model = clean_text(section.get("model", "")).upper()
            if model and int(priority) > int(existing.get("_model_priority", -1)):
                existing["model"] = model
                existing["_model_priority"] = int(priority)
            existing["pages"].update(section.get("pages", set()))
            for alias in aliases:
                if alias not in existing["aliases"]:
                    existing["aliases"].append(alias)
            if not existing.get("code") and code:
                existing["code"] = code
        for alias in existing.get("aliases", []):
            by_alias[normalize_key(alias)] = existing
        return existing

    # Printed index/catalogue entries are reliable, but an exact page header is
    # stronger and is therefore registered with a higher priority below.
    for section in index_sections:
        register(section, priority=80)

    page_metadata: dict[int, dict[str, Any]] = {}
    page_heading_sections: dict[int, list[dict[str, Any]]] = {}
    for page_number, markdown in extracted_pages:
        page = int(page_number)
        metadata = _page_metadata(page, markdown)
        page_metadata[page] = metadata
        if page in index_pages:
            continue
        if allow_simple_section_codes:
            registered_headings: list[dict[str, Any]] = []
            for heading in _classic_catalog_sections(markdown):
                registered_headings.append(
                    register(
                        {
                            "code": heading.get("section_code", ""),
                            "aliases": [heading.get("section_code", "")],
                            "name": heading.get("section_name_english", ""),
                            "maker": metadata.get("maker", ""),
                            "model": metadata.get("model", ""),
                            "pages": {page},
                        },
                        priority=110,
                    )
                )
            if registered_headings:
                page_heading_sections[page] = registered_headings
                continue
        code_tokens = catalog_code_tokens(metadata.get("section_code", ""))
        if code_tokens:
            registered = register(
                {
                    "code": _join_section_codes(code_tokens),
                    "aliases": code_tokens,
                    "name": metadata.get("section_name_english", ""),
                    "maker": metadata.get("maker", ""),
                    "model": metadata.get("model", ""),
                    "pages": {page},
                },
                priority=100,
            )
            registered["pages"].add(page)

    # AI catalogue values can fill gaps but must never overwrite an exact printed
    # header title such as RESILIENT MOUNTING with a paraphrase such as ELASTIC MOUNTING.
    #
    # In a multilingual Order-No. catalogue, an item's drawing position (for
    # example 7.15 or 15.3) looks exactly like a legitimate hierarchy code.  The
    # AI therefore must not be allowed to create a section from an unconfirmed
    # numeric value: that turns spare parts into false sub-machineries.  It may
    # enrich only a section that the source itself has already confirmed.
    for row in extracted_rows or []:
        code_tokens = catalog_code_tokens(row.get("section_code", ""))
        if not code_tokens:
            continue
        if allow_simple_section_codes:
            confirmed: list[dict[str, Any]] = []
            for token in code_tokens:
                section = by_alias.get(normalize_key(token))
                if section is not None and all(
                    id(section) != id(existing) for existing in confirmed
                ):
                    confirmed.append(section)
            if len(confirmed) != 1:
                continue
        source_page = quantity_to_number(row.get("source_page"))
        register(
            {
                "code": _join_section_codes(code_tokens),
                "aliases": code_tokens,
                "name": clean_text(
                    row.get("section_name_english", row.get("detected_machinery", ""))
                ).upper(),
                "maker": clean_text(row.get("section_maker", "")).upper(),
                "model": clean_text(row.get("section_model", "")).upper(),
                # AI rows may contain stale section context. They can define a
                # missing catalogue alias, but never claim authoritative pages.
                "pages": set(),
            },
            priority=30,
        )

    def unique_sections(values: Sequence[dict[str, Any] | None]) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        seen_ids: set[int] = set()
        for value in values:
            if value is not None and id(value) not in seen_ids:
                result.append(value)
                seen_ids.add(id(value))
        return result

    def sections_for_codes(value: Any) -> list[dict[str, Any]]:
        return unique_sections(
            [
                by_alias.get(normalize_key(token))
                for token in catalog_code_tokens(value)
            ]
        )

    def normalized_code_text(value: Any) -> str:
        text_value = clean_text(value).upper()
        return re.sub(r"\s*([./_-])\s*", r"\1", text_value)

    def alias_in_header(alias: str, header_text: str) -> bool:
        alias_text = normalized_code_text(alias)
        source_text = normalized_code_text(header_text)
        if not alias_text or not source_text:
            return False
        return bool(
            re.search(
                rf"(?<![A-Z0-9]){re.escape(alias_text)}(?![A-Z0-9])",
                source_text,
            )
        )

    parts_pages = {
        int(page)
        for page, markdown in extracted_pages
        if int(page) not in index_pages and _looks_like_spare_table_page(markdown)
    }
    page_map: dict[int, dict[str, Any]] = {}
    page_map_sources: dict[int, str] = {}
    ambiguous_pages: set[int] = set()
    unmapped_parts_pages: set[int] = set()
    active: dict[str, Any] | None = None
    active_anchor_page: int | None = None

    for page in sorted(page_metadata):
        if page in index_pages:
            continue
        metadata = page_metadata[page]
        explicit_matches = page_heading_sections.get(page) or sections_for_codes(
            metadata.get("section_code", "")
        )
        multiple_numbered_headings = bool(
            allow_simple_section_codes and len(explicit_matches) > 1
        )
        resolved: dict[str, Any] | None = None
        resolved_source = ""

        if len(explicit_matches) == 1:
            resolved = explicit_matches[0]
            resolved_source = "exact page-header code"
        elif len(explicit_matches) > 1:
            # Historical catalogues can legitimately start two numbered sections
            # on the same PDF page. Row-level parsing assigns that page's records;
            # retain the last heading only as context for the following page.
            ambiguous_pages.add(page)
            if allow_simple_section_codes:
                active = explicit_matches[-1]
                active_anchor_page = page
            else:
                active = None
                active_anchor_page = None
        else:
            header_text = metadata.get("header_text", "")
            header_matches = unique_sections(
                [
                    section
                    for section in sections
                    if any(
                        alias_in_header(alias, header_text)
                        for alias in section.get("aliases", [])
                    )
                ]
            )
            if len(header_matches) == 1:
                resolved = header_matches[0]
                resolved_source = "exact code in page header area"
            elif len(header_matches) > 1:
                ambiguous_pages.add(page)
                active = None
                active_anchor_page = None
            else:
                # Use a very strong title match only when no code is available.
                name_key = normalize_key(metadata.get("section_name_english", ""))
                if name_key:
                    scored = sorted(
                        (
                            (
                                SequenceMatcher(
                                    None,
                                    name_key,
                                    normalize_key(section.get("name", "")),
                                ).ratio(),
                                section,
                            )
                            for section in sections
                            if clean_text(section.get("name", ""))
                        ),
                        key=lambda pair: pair[0],
                        reverse=True,
                    )
                    if scored and scored[0][0] >= 0.96:
                        second_score = scored[1][0] if len(scored) > 1 else 0.0
                        if scored[0][0] - second_score >= 0.03:
                            resolved = scored[0][1]
                            resolved_source = "exact page-header title"

        if resolved is not None:
            active = resolved
            active_anchor_page = page
            resolved["pages"].add(page)
        elif (
            page in parts_pages
            and active is not None
            and active_anchor_page is not None
            and page - active_anchor_page == 1
            and page not in ambiguous_pages
            and not clean_text(metadata.get("section_code", ""))
        ):
            # Carry forward only to an immediately consecutive continuation page
            # that has no new header code. This cannot jump across ambiguous pages.
            resolved = active
            resolved_source = "confirmed consecutive continuation"
            active_anchor_page = page
            resolved["pages"].add(page)
        elif page in parts_pages and not multiple_numbered_headings:
            unmapped_parts_pages.add(page)
            active = None
            active_anchor_page = None

        if page in parts_pages and resolved is not None:
            page_map[page] = resolved
            page_map_sources[page] = resolved_source

    fallback_maker = global_maker or clean_text(main_row.get("MAKER", "")).upper()
    fallback_model = global_model or clean_text(main_row.get("MODEL", "")).upper()
    final_sections: list[dict[str, Any]] = []
    seen_section_ids: set[int] = set()
    for section in sections:
        if id(section) in seen_section_ids:
            continue
        seen_section_ids.add(id(section))
        pages = sorted(
            int(page)
            for page in section.get("pages", set())
            if int(page) not in index_pages
        )
        section["name"] = _clean_catalog_section_name(
            section.get("name", ""), section.get("code", "")
        )
        section["maker"] = clean_text(section.get("maker", "")).upper() or fallback_maker
        section["model"] = clean_text(section.get("model", "")).upper() or fallback_model
        section["first_page"] = min(pages) if pages else None
        section["last_page"] = max(pages) if pages else None
        section["pages"] = pages
        section.pop("_name_priority", None)
        section.pop("_maker_priority", None)
        section.pop("_model_priority", None)
        if section.get("code") or section.get("name"):
            final_sections.append(section)
    def section_sort_code(value: Any) -> tuple[Any, ...]:
        code = clean_text(value).strip().rstrip(".")
        if re.fullmatch(r"\d{1,2}(?:\.\d{1,2})?", code):
            return (0, *(int(part) for part in code.split(".")))
        return (1, normalize_key(code))

    final_sections.sort(
        key=lambda item: (
            item.get("first_page") is None,
            item.get("first_page") or 10**9,
            section_sort_code(item.get("code", "")),
            item.get("name", ""),
        )
    )
    return {
        "sections": final_sections,
        "page_map": page_map,
        "page_map_sources": page_map_sources,
        "page_metadata": page_metadata,
        "index_pages": index_pages,
        "parts_pages": parts_pages,
        "ambiguous_pages": ambiguous_pages,
        "unmapped_parts_pages": unmapped_parts_pages,
    }

def _catalog_prompt_hint(sections: Sequence[dict[str, Any]]) -> str:
    lines: list[str] = []
    for section in sections[:120]:
        code = clean_text(section.get("code", ""))
        name = clean_text(section.get("name", ""))
        maker = clean_text(section.get("maker", ""))
        model = clean_text(section.get("model", ""))
        pages = ",".join(str(page) for page in section.get("pages", []) if page)
        if code or name:
            lines.append(
                f"- CODE={code}; NAME={name}; PAGES={pages}; "
                f"MAKER={maker}; MODEL={model}"
            )
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


def _carry_forward_merged_item_numbers(
    rows: Sequence[dict[str, Any]],
) -> tuple[list[dict[str, Any]], int]:
    """Repeat a merged drawing position on immediate continuation records.

    Some manuals print one ITEM NO in a vertically merged cell while listing two
    or more identifiers/descriptions beside it. OCR/Markdown may therefore leave
    the continuation record's item number blank. Carry forward only within the
    same source page, section code, and table title, and only for a genuine spare
    row with its own identifier and description.
    """
    result: list[dict[str, Any]] = []
    last_key: tuple[str, str, str] | None = None
    last_item_no = ""
    inherited_count = 0

    for raw in rows:
        row = dict(raw)
        key = (
            str(row.get("source_page", "")),
            normalize_key(row.get("section_code", "")),
            normalize_key(row.get("table_title", "")),
        )
        item_no = clean_text(row.get("item_no", ""))
        identifier = clean_text(
            row.get("ident_no", row.get("code", row.get("part_no", "")))
        )
        description = clean_text(
            row.get("description_english", row.get("description", ""))
        )

        if item_no:
            last_key = key
            last_item_no = item_no
        elif (
            last_key == key
            and last_item_no
            and identifier
            and description
        ):
            row["item_no"] = last_item_no
            row["item_no_inherited"] = True
            inherited_count += 1
        else:
            # Do not carry a position across another page, section, or table.
            last_key = key
            last_item_no = ""

        result.append(row)

    return result, inherited_count


def prepare_benefit_rows(
    ai_rows: Sequence[dict[str, Any]],
    extracted_pages: Sequence[tuple[int, str]],
    source_document_name: str,
    main_row: dict[str, Any],
    default_unit: str = "PCS",
    catalog_pages: Sequence[tuple[int, str]] | None = None,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Repair and normalize extraction using exact PDF table structure and page headers."""
    normalized_ai = [_normalized_ai_row(row) for row in ai_rows if isinstance(row, dict)]
    direct_rows = _direct_table_rows(extracted_pages)
    section_context_pages = list(catalog_pages) if catalog_pages is not None else list(extracted_pages)
    catalog = build_section_catalog(section_context_pages, normalized_ai, main_row)
    sections = catalog["sections"]
    page_map = catalog["page_map"]
    page_map_sources = catalog.get("page_map_sources", {})
    index_pages = catalog["index_pages"]
    parts_pages = catalog["parts_pages"]
    ambiguous_pages = catalog.get("ambiguous_pages", set())
    unmapped_parts_pages = catalog.get("unmapped_parts_pages", set())
    allow_simple_section_codes = (
        _document_extraction_profile(section_context_pages)
        == MULTILINGUAL_ORDER_CATALOG_PROFILE
    )

    by_alias: dict[str, dict[str, Any]] = {}
    for section in sections:
        for alias in section.get("aliases", []):
            if normalize_key(alias):
                by_alias[normalize_key(alias)] = section

    def benefit_code_tokens(value: Any) -> list[str]:
        text = clean_text(value).upper().strip().rstrip(".")
        if allow_simple_section_codes:
            return (
                [text]
                if re.fullmatch(r"\d{1,2}(?:\.\d{1,2})?", text)
                else []
            )
        return _section_code_tokens(value, permissive=True)

    def exact_sections_for_code(value: Any) -> list[dict[str, Any]]:
        matches: list[dict[str, Any]] = []
        seen_ids: set[int] = set()
        for token in benefit_code_tokens(value):
            section = by_alias.get(normalize_key(token))
            if section is not None and id(section) not in seen_ids:
                matches.append(section)
                seen_ids.add(id(section))
        return matches

    def fuzzy_section_for_name(value: Any) -> dict[str, Any] | None:
        name_key = normalize_key(value)
        if not name_key:
            return None
        scored = sorted(
            (
                (
                    SequenceMatcher(
                        None,
                        name_key,
                        normalize_key(section.get("name", "")),
                    ).ratio(),
                    section,
                )
                for section in sections
                if clean_text(section.get("name", ""))
            ),
            key=lambda pair: pair[0],
            reverse=True,
        )
        if not scored or scored[0][0] < 0.94:
            return None
        second_score = scored[1][0] if len(scored) > 1 else 0.0
        return scored[0][1] if scored[0][0] - second_score >= 0.03 else None

    def section_for(
        page: int | None,
        direct_row: dict[str, Any] | None = None,
        ai_row: dict[str, Any] | None = None,
    ) -> tuple[dict[str, Any] | None, str, bool]:
        """Resolve a section with printed page context ahead of AI context."""
        direct_row = direct_row or {}
        ai_row = ai_row or {}
        direct_matches = exact_sections_for_code(direct_row.get("section_code", ""))
        page_section = page_map.get(page) if page is not None else None
        ai_matches = exact_sections_for_code(ai_row.get("section_code", ""))
        conflict = False

        if len(direct_matches) == 1:
            chosen = direct_matches[0]
            if page_section is not None and id(page_section) != id(chosen):
                conflict = True
            if len(ai_matches) == 1 and id(ai_matches[0]) != id(chosen):
                conflict = True
            return chosen, "exact printed page-header code", conflict
        if len(direct_matches) > 1:
            conflict = True

        if page_section is not None:
            if len(ai_matches) == 1 and id(ai_matches[0]) != id(page_section):
                conflict = True
            return page_section, page_map_sources.get(page, "page-header map"), conflict

        # AI section codes are only a fallback when no printed page/header mapping
        # exists. They can no longer override a deterministic source code.
        if len(ai_matches) == 1:
            return ai_matches[0], "AI section code fallback", conflict
        if len(ai_matches) > 1:
            conflict = True

        direct_name_match = fuzzy_section_for_name(direct_row.get("section_name_english", ""))
        if direct_name_match is not None:
            return direct_name_match, "printed page-title match", conflict
        ai_name_match = fuzzy_section_for_name(ai_row.get("section_name_english", ""))
        if ai_name_match is not None:
            return ai_name_match, "AI title fallback", conflict
        return None, "unconfirmed", True

    # Match AI rows with the strongest available row identity. ITEM NO alone is not
    # globally unique and could previously select a different row on the same page.
    ai_by_full_key: dict[tuple[int | None, str, str], dict[str, Any]] = {}
    ai_by_page_ident_lists: dict[tuple[int | None, str], list[dict[str, Any]]] = {}
    ai_by_page_item_lists: dict[tuple[int | None, str], list[dict[str, Any]]] = {}
    for row in normalized_ai:
        page = row.get("source_page")
        item_key = normalize_key(row.get("item_no", ""))
        ident_key = normalize_key(row.get("ident_no", ""))
        if item_key or ident_key:
            ai_by_full_key[(page, item_key, ident_key)] = row
        if ident_key:
            ai_by_page_ident_lists.setdefault((page, ident_key), []).append(row)
        if item_key:
            ai_by_page_item_lists.setdefault((page, item_key), []).append(row)

    def matched_ai_row(page: int, item_key: str, ident_key: str) -> dict[str, Any] | None:
        exact = ai_by_full_key.get((page, item_key, ident_key))
        if exact is not None:
            return exact
        by_ident = ai_by_page_ident_lists.get((page, ident_key), []) if ident_key else []
        if len(by_ident) == 1:
            return by_ident[0]
        by_item = ai_by_page_item_lists.get((page, item_key), []) if item_key else []
        if len(by_item) == 1:
            return by_item[0]
        return None

    output: list[dict[str, Any]] = []
    matched_ai_ids: set[int] = set()
    language_exceptions = 0
    paraphrases_prevented = 0
    section_conflicts = 0
    unconfirmed_sections = 0

    if direct_rows:
        for direct in direct_rows:
            page = int(direct["source_page"])
            item_key = normalize_key(direct.get("item_no", ""))
            ident_key = normalize_key(direct.get("ident_no", ""))
            ai = matched_ai_row(page, item_key, ident_key)
            if ai is not None:
                matched_ai_ids.add(id(ai))

            ai_description_source = ""
            if ai is not None and not bool(ai.get("local_fallback", False)):
                ai_description_source = ai.get("description_english", "")
            description, language_review, ai_disagreed, description_source = (
                _select_source_faithful_description(
                    direct.get("description_raw", ""),
                    ai_description_source,
                )
            )
            if language_review:
                language_exceptions += 1
            if ai_disagreed:
                paraphrases_prevented += 1

            section, section_source, section_conflict = section_for(page, direct, ai)
            if section_conflict:
                section_conflicts += 1
            if section is None:
                unconfirmed_sections += 1

            # Printed page data wins over AI fallbacks for section identity and title.
            section_code = (
                clean_text((section or {}).get("code", ""))
                or clean_text(direct.get("section_code", ""))
                or clean_text((ai or {}).get("section_code", ""))
            )
            section_name = (
                clean_text((section or {}).get("name", ""))
                or clean_text(direct.get("section_name_english", ""))
                or clean_text((ai or {}).get("section_name_english", ""))
            )
            section_maker = (
                clean_text((section or {}).get("maker", ""))
                or clean_text(direct.get("section_maker", ""))
                or clean_text((ai or {}).get("section_maker", ""))
            )
            section_model = (
                clean_text((section or {}).get("model", ""))
                or clean_text(direct.get("section_model", ""))
                or clean_text((ai or {}).get("section_model", ""))
            )
            first_page = (section or {}).get("first_page") or page

            confidence = max(0.88, clamp_confidence(direct.get("confidence", 0.88)))
            if description_source != "printed English":
                confidence = min(
                    confidence,
                    clamp_confidence((ai or {}).get("confidence", 0.78), 0.78),
                    0.82,
                )
            if language_review:
                confidence = min(confidence, 0.60)
            if section_conflict:
                confidence = min(confidence, 0.62)
            if section is None:
                confidence = min(confidence, 0.50)

            ident_no = clean_text(direct.get("ident_no", ""))
            resolved_item_no = clean_text(direct.get("item_no", "")) or clean_text(
                (ai or {}).get("item_no", "")
            )
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
                    "item_no": resolved_item_no,
                    "description_english": description.upper(),
                    "description": description.upper(),
                    "source_description_raw": _clean_markdown_cell(direct.get("description_raw", "")),
                    "source_section_name_raw": _clean_markdown_cell(direct.get("table_title", "")),
                    "unit": normalize_unit((ai or {}).get("unit", ""), default_unit),
                    "quantity": quantity_to_number(direct.get("quantity")),
                    "confidence": confidence,
                    "language_review": language_review,
                    "description_source": description_source,
                    "description_source_mismatch": ai_disagreed,
                    "section_review": bool(section_conflict or section is None),
                    "section_assignment_source": section_source,
                }
            )
    else:
        # Irregular/non-Markdown manuals still use the strict AI schema, but section
        # assignment remains anchored to the page map whenever one exists.
        for ai in normalized_ai:
            page = ai.get("source_page")
            if page in index_pages:
                continue
            section, section_source, section_conflict = section_for(page, None, ai)
            ident_no = clean_text(ai.get("ident_no", ""))
            description, language_review = _best_effort_english_description(
                ai.get("description_english", "")
            )
            if not description or not (ident_no or clean_text(ai.get("item_no", ""))):
                continue
            confidence = clamp_confidence(ai.get("confidence", 0.70))
            if language_review:
                confidence = min(confidence, 0.60)
                language_exceptions += 1
            if section_conflict:
                confidence = min(confidence, 0.62)
                section_conflicts += 1
            if section is None:
                confidence = min(confidence, 0.50)
                unconfirmed_sections += 1
            output.append(
                {
                    **ai,
                    "section_code": clean_text((section or {}).get("code", ai.get("section_code", ""))),
                    "section_name_english": clean_text((section or {}).get("name", ai.get("section_name_english", ""))).upper(),
                    "detected_machinery": clean_text((section or {}).get("name", ai.get("section_name_english", ""))).upper(),
                    "section_maker": clean_text((section or {}).get("maker", "")) or clean_text(ai.get("section_maker", "")),
                    "section_model": clean_text((section or {}).get("model", "")) or clean_text(ai.get("section_model", "")),
                    "section_start_page": (section or {}).get("first_page") or ai.get("section_start_page") or page,
                    "part_no": ident_no,
                    "code": ident_no,
                    "description_english": description,
                    "description": description,
                    "source_description_raw": clean_text(ai.get("description_english", "")),
                    "source_section_name_raw": clean_text(ai.get("section_name_english", "")),
                    "confidence": confidence,
                    "language_review": language_review,
                    "description_source": "AI English/translation",
                    "description_source_mismatch": False,
                    "section_review": bool(section_conflict or section is None),
                    "section_assignment_source": section_source,
                }
            )

    # AI-only rows are accepted only on pages that contain an actual parts table.
    if direct_rows:
        existing_keys = {
            (
                int(row.get("source_page") or 0),
                normalize_key(row.get("item_no", "")),
                normalize_key(row.get("ident_no", "")),
            )
            for row in output
        }
        for ai in normalized_ai:
            if id(ai) in matched_ai_ids or ai.get("source_page") not in parts_pages:
                continue
            key = (
                int(ai.get("source_page") or 0),
                normalize_key(ai.get("item_no", "")),
                normalize_key(ai.get("ident_no", "")),
            )
            if key in existing_keys:
                continue
            description, language_review = _best_effort_english_description(
                ai.get("description_english", "")
            )
            if description and (clean_text(ai.get("ident_no", "")) or clean_text(ai.get("item_no", ""))):
                section, section_source, section_conflict = section_for(
                    ai.get("source_page"), None, ai
                )
                ident_no = clean_text(ai.get("ident_no", ""))
                confidence = min(clamp_confidence(ai.get("confidence", 0.70)), 0.82)
                if language_review:
                    confidence = min(confidence, 0.60)
                    language_exceptions += 1
                if section_conflict:
                    confidence = min(confidence, 0.62)
                    section_conflicts += 1
                if section is None:
                    confidence = min(confidence, 0.50)
                    unconfirmed_sections += 1
                output.append(
                    {
                        **ai,
                        "section_code": clean_text((section or {}).get("code", ai.get("section_code", ""))),
                        "section_name_english": clean_text((section or {}).get("name", ai.get("section_name_english", ""))).upper(),
                        "detected_machinery": clean_text((section or {}).get("name", ai.get("section_name_english", ""))).upper(),
                        "section_maker": clean_text((section or {}).get("maker", "")) or clean_text(ai.get("section_maker", "")),
                        "section_model": clean_text((section or {}).get("model", "")) or clean_text(ai.get("section_model", "")),
                        "section_start_page": (section or {}).get("first_page") or ai.get("source_page"),
                        "part_no": ident_no,
                        "code": ident_no,
                        "description_english": description,
                        "description": description,
                        "source_description_raw": clean_text(ai.get("description_english", "")),
                        "source_section_name_raw": clean_text(ai.get("section_name_english", "")),
                        "confidence": confidence,
                        "language_review": language_review,
                        "description_source": "AI-only row",
                        "description_source_mismatch": False,
                        "section_review": bool(section_conflict or section is None),
                        "section_assignment_source": section_source,
                    }
                )

    output, inherited_item_count = _carry_forward_merged_item_numbers(output)

    # Deterministic de-duplication by source page + item + identifier. PART NO alone
    # is intentionally not treated as globally unique across a technical manual.
    deduplicated: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str]] = set()
    for row in output:
        identifier_key = normalize_key(row.get("ident_no", row.get("code", "")))
        if allow_simple_section_codes and identifier_key:
            # Order-No. is the unique catalogue identifier. If the same code is
            # printed again for another applicability line, keep one import row.
            key = ("ORDERNO", "", identifier_key)
        else:
            key = (
                str(row.get("source_page", "")),
                normalize_key(row.get("item_no", "")),
                identifier_key,
            )
        if key in seen:
            continue
        seen.add(key)
        deduplicated.append(row)

    messages = [
        f"Deterministic table verification found {len(direct_rows)} spare-part row(s).",
        f"Detected {len(sections)} source-coded sub-machinery section(s).",
    ]
    if inherited_item_count:
        messages.append(
            f"Repeated {inherited_item_count} merged/continuation ITEM NO value(s) on linked spare-part rows."
        )
    if paraphrases_prevented:
        messages.append(
            f"Kept the exact printed English wording for {paraphrases_prevented} row(s) where the AI proposed a different phrase."
        )
    if language_exceptions:
        messages.append(
            f"{language_exceptions} description(s) could not be confidently isolated as English and remain low-confidence review items."
        )
    if section_conflicts:
        messages.append(
            f"Resolved {section_conflicts} section conflict(s) by giving the printed page header priority over AI context."
        )
    if ambiguous_pages:
        messages.append(
            "Ambiguous section headers detected on PDF page(s): "
            + ", ".join(str(page) for page in sorted(ambiguous_pages))
            + ". Previous section context was not carried across those pages."
        )
    if unmapped_parts_pages:
        messages.append(
            "No confirmed section mapping was available for parts-table page(s): "
            + ", ".join(str(page) for page in sorted(unmapped_parts_pages))
            + ". Those rows were kept at low confidence instead of inheriting a stale section."
        )
    if unconfirmed_sections:
        messages.append(
            f"{unconfirmed_sections} row(s) remain without a fully confirmed source section."
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
        name = _clean_catalog_section_name(
            _clean_machinery_name(
                row.get("DETECTED MACHINERY", row.get("TABLE TITLE", ""))
            ),
            row.get("SECTION CODE", ""),
        )
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
        benefit_name = _clean_catalog_section_name(item["name"], item["code"])
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
        source_warnings: list[str] = []
        if bool(raw.get("description_source_mismatch", False)):
            source_warnings.append(
                "Source: AI wording differed - exact printed English was kept"
            )
        if bool(raw.get("language_review", False)):
            source_warnings.append("Source: English wording was not fully confirmed")
        section_assignment_source = clean_text(raw.get("section_assignment_source", ""))
        if bool(raw.get("section_review", False)):
            if section_assignment_source == "unconfirmed":
                source_warnings.append("Source: section could not be confirmed")
            else:
                source_warnings.append(
                    "Source: AI/page section conflict - printed page header was kept"
                )
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
                "WARNING": "; ".join(source_warnings),
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
        existing_warning = clean_text(row.get("WARNING", ""))
        row_messages: list[str] = [
            message.strip()
            for message in existing_warning.split(";")
            if message.strip().startswith("Source:")
        ]
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
