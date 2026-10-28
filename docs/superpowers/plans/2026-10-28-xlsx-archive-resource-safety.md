# Excel Input Archive Safety Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans task-by-task, with superpowers:test-driven-development and verification-before-completion.

**Goal:** Prevent excessive memory consumption from pathological Excel ZIP members without changing normal merge results.

**Architecture:** A new xlsx_safety.validate_xlsx_archive(path: Path) function checks ZIP metadata immediately before load_workbook. Existing code continues unchanged for ordinary ZIP packages; a failure follows the existing SourceFileError and GUI guidance path.

**Tech Stack:** Python 3.12, zipfile, openpyxl 3.1.5, unittest, GitHub Windows Actions.

**Spec:** docs/superpowers/specs/2026-10-28-xlsx-archive-resource-safety.md

## Global Constraints
- Preserve existing CSV, XLS, XLSX, XLSM conversion, time-named sheets, text/formula/style types, checkbox selection, output numbering and UI.
- Preserve source and previously generated XLSX; clean owned scratch.
- No new third-party dependencies, no security bypass.
- The 8-hour company Windows soak remains last.

## Review Focus
- Extremely compressible syntactically valid worksheet -> preflight blocks before parsing.
- Small highly compressible worksheet -> normal output.
- Duplicated xl/worksheets member or [Content_Types].xml -> clear rejection.
- Impossibly large central directory aggregate/XML part -> rejection without output.
- Existing result when unsafe new source is rejected -> untouched, no partial new XLSX.

### Task 1: Red tests
**Files:** tests/test_xlsx_archive_safety.py
- [x] Create valid 64+ MiB expanded XML fixture and test failure before parser invocation. Before fix: SourceFileError not raised.
- [x] Test normal small compressed XLSX and existing file preservation.
- [x] Add tests for duplicate ZIP manifest and component metadata size limits. Before manifest fix: SourceFileError not raised.

### Task 2: Guard and Korean diagnostic
**Files:** src/excel_sheet_merger/xlsx_safety.py, src/excel_sheet_merger/merge.py, src/excel_sheet_merger/failure_guidance.py
- [x] Implement central-directory size/ratio/part-count/duplicate checks without inflating source bytes.
- [x] Invoke preflight only before openpyxl workbook parsing for XLSX/XLSM.
- [x] Add specific EXCEL_RESOURCE_LIMIT diagnostic (RED: generic EXCEL_OPEN, GREEN: dedicated guidance).
- [x] Run targeted tests and Linux full suite: 137 tests, OK, 4 Windows-only skipped.

### Task 3: Verification and publication
**Files:** README.md, RELEASE_NOTES.md, docs/superpowers/stability-recheck-2026-10-28-round9.md
- [x] Document boundary/legitimate-outlier tradeoff, and record before/after behavior.
- [x] Run Python compile and Git diff check; test suite green.
- [ ] Publish only exact tested source/tests and read-current documentation onto GitHub main (no other files).
- [ ] Verify Windows CI, built EXE smoke and released assets against the exact new commit. Do not claim success without evidence.
