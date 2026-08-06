# Spare Parts OCR Import Builder — catalogue hierarchy update 4.9.1

Replace both deployed `app.py` and `tools.py` with the accompanying revised files.
Keep `vessels.csv`, the Excel template, Streamlit secrets, and requirements in the
same locations.

## What changed

### 4.9.1 catalogue extraction corrections

- Detects standalone major headings such as `1. BURNER CASTING AND INDIVIDUAL
  PARTS` and `3. SERVO DRIVE FOR OIL BURNERS`, even when OCR places the heading
  above rather than inside the Markdown table.
- Stops promoting repeated item positions, Order-No. values, or spare descriptions
  to sub-machineries. Decimal sub-machinery codes must now satisfy stronger
  hierarchy evidence.
- Treats German assembly headings such as `Brennermotor`, `Luftregelung`, and
  `Stellantrieb` as title text, never as the maker. The main-machinery maker is used
  when the source does not print a section-specific manufacturer.
- Keeps sub-machinery names clean: no appended `(code)`, ordering note, repeated
  title, or first-spare description.
- Splits flattened Order-No. cells into one unique spare-part row per printed code.
- Measures multilingual catalogue coverage against deterministic table evidence.
  Suspiciously sparse pages are retried once and, when adaptive analysis is enabled,
  selectively escalated to the configured high-accuracy model. A failed recovery
  never discards rows already extracted by the normal model.
- Adds regression coverage for all of the above GSL Tegea failure modes.

### Earlier 4.9.0 workflow and adaptive-analysis improvements

- Removed the fixed-height workflow scroll surface. Each step now uses the normal
  browser page scroll.
- Restored Streamlit Cloud's page-level scroll surface so review tables can extend
  below the initial browser viewport.
- Added a sticky, status-aware workflow navigator with completion marks, issue
  counts, progress, and contextual Back/Continue actions.
- Moved OCR tuning, API-key fallback, template replacement, and reset controls into
  one collapsed **Processing & export settings** panel.
- Added explicit confirmation before removing a document or resetting its OCR and
  review state.
- Fixed export helpers being scoped to the Review step, which could cause
  `name 'export_filename' is not defined` when opening Export directly.
- Replaced form-based review saving with per-cell autosave for sub-machinery and
  spare-part editors.
- Automatically recalculates readiness, invalidates stale exports, and persists the
  active document after review edits.
- Reduced the default spare-parts review grid to the fields needed for correction;
  technical matching fields remain in the audit workbook.
- Moved the build number into About and replaced the main caption with a user-facing
  product description.
- Added automatic detection for historical German/English/French catalogues headed
  `Bestell-Nr. / Order-No. / No de commande`.
- Maps every Order-No. to both PART NO and CODE, maps Bild/Pict./Photo to ITEM NO,
  and expands vertically merged variant rows into one spare row per Order-No.
- Selects the English Designation column deterministically and prevents the adjacent
  German/French columns from replacing it.
- Treats `ca. kg / appr. kg / env. kg` as weight, never as QNT.
- Detects simple numbered catalogue sections inside table rows, carries them through
  continuation pages, and handles two sections beginning on the same page.
- Retries dense catalogue pages automatically when the AI returns valid but empty
  JSON, then supplements extraction from the OCR table itself.
- Applies strict global Order-No. de-duplication for this catalogue profile.
- Recognizes both major numeric sections (`1`, `3`, `4`) and printed decimal
  subsections (`3.30`, `6.40`) while preserving trailing zeroes.
- Uses the structural Item/Photo and Order-No. columns to distinguish hierarchy
  headings from ordinary positions, instead of accepting model names such as
  `SQM10` or `ASZ12` as sub-machinery codes.
- Removes merged child descriptions and ordering notes from sub-machinery names;
  for example, `BURNER MOTOR` and `BLOWER` remain clean headings.
- Preserves the first spare row when OCR merges it into the same table row as a
  section heading.
- Orders sub-machineries by their printed numeric hierarchy when several begin on
  the same page, rather than alphabetically by name.
- Adds an enabled-by-default adaptive document-analysis pass using
  `mistral-large-2512` and the existing `MISTRAL_API_KEY`.
- Samples at most 24 evenly distributed OCR pages, then records the document's
  languages, English-selection rule, part/item column meanings, hierarchy patterns,
  continuation behavior, exclusions, uncertainties, and profile confidence.
- Saves the resulting profile with each document and supplies it to every
  small-model extraction batch as evidence-based guidance.
- Uses Large 3 only for language terms that remain uncertain after deterministic
  English selection; routine row extraction remains on `mistral-small-latest`.
- Displays the saved profile in the Processing step and falls back to the previous
  extraction path if Large 3 is unavailable, without losing vessel or machinery data.

## Suggested acceptance check

1. Upload a small PDF and complete the Machinery step.
2. Confirm the Continue button activates and opens Processing.
3. Run OCR and confirm the stepper shows completion/issue states.
4. Edit one sub-machinery cell, leave the cell, and confirm the autosave message.
5. Edit and manually verify a spare row, then change pages and return; confirm the
   edit remains.
6. Confirm Export remains blocked while included rows need correction and becomes
   available when all included rows are ready.
7. Confirm document removal and OCR reset require their confirmation checkbox.
8. For the GSL Tegea catalogue, confirm item `3.82` creates three rows with codes
   `151 518 1508/2`, `151 707 1503/2`, and `151 907 1505/2`, all under sub-machinery
   `SERVO DRIVE FOR OIL BURNERS` with CODE `3` stored in its separate column.
9. Confirm the catalogue includes `BURNER CASING AND INDIVIDUAL PARTS`,
   `SERVO-DRIVE FOR GAS AND DUAL FUEL BURNERS`, and `MAGNET COUPLING`, with
   codes `1`, `3.30`, and `6.40` stored only in CODE; confirm sections 4 and 5 are named only
   `BURNER MOTOR` and `BLOWER`.
10. In Processing settings, confirm adaptive document analysis is On and the model
    is `mistral-large-2512`. Run OCR, then open **Adaptive document analysis** and
    verify the displayed languages, column roles, hierarchy examples, and confidence.
11. Temporarily enter an unavailable analysis model and confirm OCR continues through
    the safe fallback path with an informational recovery message.

Syntax compilation, isolated autosave callback tests, multilingual Order-No.
catalogue parser/section tests, bounded profile-sampling tests, profile normalization,
and extraction-prompt integration tests passed in the Codex runtime. A live Mistral
API call was not performed because no API credential was available in this workspace.
