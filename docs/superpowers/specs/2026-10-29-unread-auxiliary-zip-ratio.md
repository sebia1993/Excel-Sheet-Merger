# Ignored Excel Custom XML ZIP Ratio — 14th Stability Design

**Date:** 2026-10-29. **Goal:** accept workbooks with high-compression Office auxiliary XML that pinned openpyxl 3.1.5 does not read, without relaxing safeguards for parsed workbook parts.

## Reproduction
Previously `validate_xlsx_archive` applies its high-ratio rejection to *every* ZIP part. A workbook containing data and `customXml/item1.xml` with 160,013 expanded bytes / 186 compressed bytes (>860:1) fails under a reduced test-only member-size threshold, while `openpyxl.load_workbook` correctly reads the sheet and ignores the auxiliary part. A real 69MiB customXml member / ~70KiB compressed also reproduces the scenario under the production 64MiB threshold.

## Selected change
Only for exact conventional `customXml/item[1-9][0-9]*.xml` part names, exempt the high-ratio heuristic. Keep MAX_ARCHIVE_MEMBERS (30,000), MAX_EXPANDED_PACKAGE_BYTES (1GiB), MAX_EXPANDED_XML_BYTES (256MiB), and duplicate-member checks, plus all high-ratio checks for other parts. The ignored custom XML is *not retained* in the result, as before.

Reject alternatives: globally raise compression ratio or uncompressed-size limits; parse massive worksheet XML merely to guess if safe.

## Invariants
No changes to CSV/XLSX/XLSM/XLS inputs, group checkbox UI, time sheets, formulas/styles/comments, source and preexisting output protection, numbered result names, failure UI, or EXE build requirements. High-ratio actual worksheet XML must be blocked before openpyxl; arbitrary customXml names must stay blocked. No background or 8-hour Windows soak in this step.

## Residual risk
Active worksheets with extraordinarily compressible legitimate cell text still may be conservatively refused. Reassess auxiliary exemption on openpyxl version changes, because it relies on the pinned parser not reading customXml item parts.
