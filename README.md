# Spare Parts OCR Import Builder — durable Step-3 review decisions 4.18.7

Replace both deployed `app.py` and `tools.py` with the accompanying revised files.
Keep `vessels.csv`, the Excel template, Streamlit secrets, and requirements in the
same locations.

## What changed

### 4.18.7 durable Step-3 review decisions

- Saves the complete submitted sub-machinery table instead of reading Streamlit's
  internal `edited_rows` callback delta, eliminating a form-submit race that could
  discard a checkbox change.
- Persists excluded proposal identities in a per-document decision ledger and
  reapplies them after every automatic title/code/evidence refresh.
- Lets an explicit re-selection or **Verify all sub-machineries** clear the saved
  exclusion, while a fresh non-append OCR run intentionally starts a new review.

### 4.18.6 title-box hierarchy correction

- Prefers a compact labelled drawing-title value over the broader page/chapter
  heading when both refer to the same source page and exact document code. For MSC
  ROMA page 358, code `9007280` is therefore named `CABLE`, not `UV REACTOR AND
  LDC / LAMP POWER CABLE`.
- Rejects OCR codes contaminated by pagination markers such as
  `9025986_Page_5` / `9025986PAPER5`; connection-list labels such as `AIR (TOTAL
  SYSTEM CONSUMPTION)` can no longer become sub-machineries.
- Removes historical excluded zero-part page-marker artifacts from saved review
  sessions and keeps future INCLUDE-only decisions separate from identity edits.

### 4.18.5 persistent sub-machinery exclusions

- Keeps an unchecked Step 3 `INCLUDE` decision after **Save sub-machinery
  changes**, subsequent Streamlit reruns, and automatic source-evidence refreshes.
- Treats inclusion/exclusion as a user review decision while continuing to refresh
  OCR-derived names, codes, pages, confidence, and linked-part counts.
- Continues to remove stale automatic candidates and preserve manual rows; an
  excluded candidate is not recreated as selected merely because it is detected
  again in the source PDF.

### 4.18.4 drawing title-block names and assembly spares

- Reconciles the native PDF chapter heading with OCR title-block evidence by exact
  page and document number. The compact labelled drawing **Title** wins when OCR
  confirms the same code with stronger evidence; for MSC ROMA document `9007280`,
  the sub-machinery NAME is therefore `CABLE`, not the broader chapter heading.
- Converts every valid equipment drawing into one linked whole-assembly spare row
  using its Document No. as PART NO/CODE and its title-block Title as DESCRIPTION.
  Drawings such as page 333 `9024508 / FILTER` are consequently represented on both
  import sheets rather than becoming an orphan candidate.
- Reconciles an already extracted exact-code assembly row instead of creating a
  duplicate, preserving the strict unique-code rule.
- Upgrades saved OCR sessions locally when Step 3 opens; no additional AI call is
  needed to add the drawing-assembly rows or apply the title-block name.
- Extends the MSC ROMA regression to require `CABLE`, all 22 drawing parents, all 22
  assembly spares, and the previously recovered maintenance/cable parts.

### 4.18.3 strict Main → Sub-machinery → Spare hierarchy

- Exports a sub-machinery only when at least one included spare-part row uses its
  exact approved name. Unlinked equipment/drawing candidates remain visible in
  Step 3 and in the audit workbook.
- Requires every included spare row to belong to an included sub-machinery; a
  main-machinery fallback can no longer pass final review or export validation.
- Blocks empty-spare exports, missing parent links, orphan sub-machineries, and any
  result where the exported sub-machinery count exceeds the spare-part count.
- Shows **Export-linked** rather than the broader included-candidate count in Step 3
  and displays the final hierarchy counts before Create Excel.
- Adds a regression for MSC ROMA: 22 detected drawing candidates remain auditable,
  while exactly 5 linked sub-machineries and 16 spare rows reach the import workbook.

### 4.18.2 authoritative source refresh and spare ownership

- Keeps OCR and native searchable-PDF evidence in separate layers. A native title
  block now replaces conflicting OCR machinery candidates from the same page
  instead of allowing both versions to survive.
- Rebuilds automatic sub-machinery proposals from current evidence on every OCR
  run and migration pass. Stale automatic rows are removed, while user-created and
  manually edited rows remain preserved.
- Rejects warning sentences, generic page labels, page-suffixed descriptions, and
  other prose fragments as automatic sub-machinery names.
- Treats explicit and recommended spare-number tables as authoritative source rows:
  damaged OCR duplicates are replaced, omitted source rows are restored, and the
  printed module ownership, quantity, and unit take precedence.
- Adds an MSC ROMA regression covering all 22 drawing sub-machineries and all 16
  source-backed maintenance/cable spare rows, including CIP MODULE and FILTER
  ownership.

### 4.18.1 BWTS hierarchy and maintenance-spares correction

- Uses the PDF's searchable text layer as independent evidence alongside Mistral
  OCR. This restores drawing titles and document numbers when OCR damages rotated
  title blocks, while OCR remains responsible for spatial table reconstruction.
- Normalizes source-layout artifacts such as `90121 03` to drawing code `9012103`
  and separates a drawing revision from codes such as `590066 1`.
- Lets authoritative equipment drawings replace weaker spare-heading proposals,
  preventing spare descriptions and generic page labels from becoming machinery.
- Retains genuine semantic parents that have no printed drawing code by assigning a
  stable `SUB-###` review code; the parent is no longer discarded before Step 3.
- Recovers maintenance-list and recommended-spares rows, including bare procurement
  labels such as `Part number: 596250 01`, their quantity, unit, and owning module.
- Keeps strict export uniqueness while allowing duplicated source evidence to be
  reconciled during review rather than blocked prematurely.

### 4.11.0 source hierarchy and code-aware sorting

- Assigns numbered-catalogue spare rows from their printed drawing-position
  hierarchy when a page contains multiple sections: for example, `7.15` belongs
  to heading `7`, while `3.31` belongs to heading `3.30`.
- Retains a recovered final zero in confirmed decimal heading codes such as
  `6.40`, rather than exporting it as `6.4`. Stored codes remain text because
  their exact formatting is business data.
- Adds **Section code** to the Review sort menu and uses natural code sorting for
  both it and **Part number**. Thus `1, 2, 10, 11` appears in that order without
  coercing codes into Excel numbers.

### 4.10.0 optional OpenAI cross-check

- Adds one optional OpenAI Responses API document-pattern verification pass after
  Mistral OCR, using `OPENAI_API_KEY` from Streamlit Secrets and
  `gpt-5.6-terra` as the configurable default model.
- Merges evidence-based OpenAI hierarchy/language guidance into the existing
  Mistral profile while keeping Mistral OCR, extraction, and local parsing as the
  primary workflow.
- Fails open for missing or expired keys, restricted models, exhausted quota,
  rate limits, network errors, timeouts, and malformed replies. The app logs the
  bypass and continues normally with Mistral.
- Performs one bounded OpenAI request with no automatic retry, reducing accidental
  usage on limited projects.
- Redacts the OpenAI key from error messages and never offers a browser/UI key field;
  the key must remain server-side in Streamlit Secrets.
- Keeps deployment compatible if `app.py` is refreshed before `tools.py`: the
  optional verifier is bypassed instead of causing an import failure.
- Aligns Sub-machineries rows/page choices with Review: 10, 25, and 50.

### 4.9.3 hierarchy gate

- Prevents AI item positions such as `7.15` and `15.3` from creating false
  sub-machineries unless the source PDF has independently confirmed the code as a
  printed hierarchy heading.

### 4.9.2 deployment compatibility

- Adds a compatibility check before passing the sparse-page recovery option to the
  extraction helper. This prevents processing from failing if a deployment briefly
  contains a newer `app.py` with an older `tools.py`; the normal extraction path
  continues and the app reports that the optional high-accuracy fallback was skipped.

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
