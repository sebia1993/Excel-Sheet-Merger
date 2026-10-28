# Excel + CSV Mixed Source Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans or subagent-driven-development. Test first then implement.

**Goal:** Extend Excel Sheet Merger with CSV+XLSX/XLSM/XLS mixed inputs; preserve existing GUI and CSV processing.
**Stack:** Python 3.12; openpyxl; xlrd; Tkinter; ttkbootstrap; PyInstaller.
**Spec:** docs/superpowers/specs/2026-10-28-excel-input-support-design.md

## Global Constraints
- No change to existing CSV semantics or user data safety.
- Excel input has exactly one worksheet; mixed groups retain filename-time sorted sheets.
- Excel formulas, styles, merges and dimensions retained where supported; VBA, charts and external features are not guaranteed.

## Review Focus
- Duplicate hours with heterogeneous source types.
- CSV formula-like text and leading zeros in a mixed group.
- Corrupt source/2-sheet workbook leaves no partial output.
- XLS importer included in executable.
- GUI selection and worker error logging unchanged.

### Task 1 — Filename classification
- [x] Add tests for mixed group, accepted extensions and rejected xlsb.
- [x] Verify RED; implement allowed extensions; verify GREEN.
Files: classify.py, tests/test_classify.py.

### Task 2 — Merge engine
- [x] Add tests for CSV+Excel, cell styles/types/formulas, duplicate hour, XLSM, XLS mock, error isolation and no-overwrite.
- [x] Verify RED; implement CSV streaming fast path and Excel normal workbook branch; verify GREEN.
Files: merge.py, tests/test_excel_inputs.py.

### Task 3 — UI and packaging
- [x] Add tests for GUI labels and dependency inclusion.
- [x] Verify RED; update UI strings, requirements, PyInstaller and workflow; verify GREEN.
Files: gui.py, UI/distribution tests, requirements-build.txt, build_windows.bat, .github/workflows/windows-exe.yml.

### Task 4 — Release
- [x] Local 43 test suite, compileall and diff hygiene.
- [ ] Verify GitHub blobs match tested source, run Windows CI and confirm EXE+SHA256 Release assets.
