# CSV + Excel Mixed Inputs — Design Spec

Date: 2026-10-28. Status: implementation requested and approved.

## Goal
Support .csv, .xlsx, .xlsm, and .xls inputs in the existing Windows 11 Excel Sheet Merger. Files with a shared site/date become one .xlsx with filename-time tabs (11시, 14시, 17시). Mixed source formats must work.

## Decisions
- Reuse classifier, sorting, duplicate-hour naming, preview selection, log, and no-overwrite semantics.
- CSV values stay as text, including leading zeros and formula-looking strings, using the existing streaming writer for CSV-only groups.
- Groups with Excel input use normal openpyxl workbook to copy common cell styles, date/number types, formulas, merged ranges, dimensions and freeze panes.
- .xlsm macros and .xls legacy formatting are not copied; .xls is read with xlrd. Images, charts, conditional formatting, pivots, external connections and cached formula values are not guaranteed.
- Exactly one sheet in each Excel source. Multi-sheet sources fail their group rather than losing data.
- .xlsb and unrecognized names are excluded with reason. Original files never modified, existing results never overwritten, temporary files cleaned after failure, other groups continue.
- ttkbootstrap office GUI and background worker architecture remain intact with updated labels. No Microsoft Excel or Python installation required for packaged EXE.

## Verification
- Test-first: mixed CSV+XLSX data, styles/types/merges, Excel-only, XLSM, duplicate hours, corrupt Excel, multiple sheets, existing outputs, XLS legacy mock reader.
- Full CSV and UI regression, Python compilation, git diff checks.
- Windows GitHub Actions bundles xlrd and releases EXE and checksum under v1.3.
- User Windows workplace verification remains separate from CI.
