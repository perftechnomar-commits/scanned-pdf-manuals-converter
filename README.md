Spare Parts OCR Import Builder — Build 4.7

Main purpose
============
Build 4.7 addresses multilingual OCR output such as:
- KURBELGEHÄUSE CRANK CASE CARTER
- WELLENDICHTRING RADIAL PACKING RING BAGUE D'ÉTANCHÉITÉ
- MANOMETER 1ST STAGE

English handling
================
1. A mandatory English-only normalization pass now runs after row matching and before review/export.
2. When German / English / French are printed together, the app isolates only the printed English wording.
3. When OCR has flattened the three language lines into one string, the cleanup pass identifies the English span rather than exporting all languages together.
4. When no English wording is present, the full technical term is translated into English.
5. Correct printed English descriptions are not retranslated or paraphrased.
6. Unique sub-machinery titles are normalized once per section code and propagated consistently to all linked spare rows.
7. Source description and source section-title text are preserved internally for the English cleanup pass.
8. The extraction prompt now explicitly forbids multilingual concatenation and includes examples from the reported manual.
9. German/French detection was expanded and the previous permissive “any ASCII text is probably English” fallback was removed.
10. Unresolved language cases remain low-confidence review items instead of silently passing as English.

Section matching retained from Build 4.6
========================================
- Section/sub-machinery assignment remains anchored to the drawing/table code printed in each page header.
- Codes in spare rows, descriptions, and cross-references cannot change the active section.
- Previous section context is carried only to a confirmed consecutive continuation page.
- Printed page-header context has priority over stale AI context.

Validation performed
====================
- Python syntax compilation passed for app.py and tools.py.
- Module import test passed.
- Synthetic English normalization tests passed for:
  * KURBELGEHÄUSE CRANK CASE CARTER -> CRANK CASE
  * multilingual GLEITLAGER / SLIDE BEARING / PALIER -> SLIDE BEARING
  * one canonical English title propagated to all rows with the same section code
  * already confirmed printed English not being unnecessarily replaced
- Synthetic section regression test passed for a stale previous AI section code being overruled by a new printed page-header code.

Deployment note
===============
Replace both app.py and tools.py together. Restart/redeploy Streamlit and rerun OCR/structuring for the affected PDF. Existing review tables and Excel exports were created with the previous language-normalization logic and will not update automatically.
