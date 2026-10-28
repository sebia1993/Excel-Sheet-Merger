# Ignored Auxiliary XML ZIP False Positive Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans and test-driven-development, task-by-task.

**Goal:** prevent false rejection of workbooks whose *unread* Office custom XML part alone has an extreme compression ratio.
**Architecture:** add a narrow exact-name predicate at the existing ratio check; leave all other ZIP limits untouched.
**Tech stack:** Python 3.12, openpyxl 3.1.5, unittest, Windows GitHub Actions/PyInstaller.
**Spec:** `docs/superpowers/specs/2026-10-29-unread-auxiliary-zip-ratio.md`.

## Global constraints
Preserve all normal merge outputs, GUI and names. Defer real-company Windows 8-hour test.

## Review focus
1. Real workbook + high-ratio customXml/item1.xml still merges (CSV ID / Excel formulas intact).
2. High-ratio xl/worksheets/sheet1.xml remains blocked before parser runs.
3. Nonstandard customXml/other.xml remains blocked.
4. Too-large ignored customXml is still rejected by XML byte cap.
5. Source files and preexisting output are preserved, with unchanged workbook semantics.

## Task 1: TDD and implementation
**Files:** `src/excel_sheet_merger/xlsx_safety.py`, `tests/test_xlsx_archive_safety.py`.
- [x] Write four real ZIP regression tests, run on unchanged logic and observe expected false-positive failure.
- [x] Add only exact `customXml/itemN.xml` exemption at the high-compression-ratio predicate.
- [x] Re-run targeted 13 tests: all pass.
- [x] Run full Linux/Xvfb suite: 161 run, 4 Windows-only skips, no errors; compileall passes.
- [x] Run real 69MiB auxiliary ZIP fixture through `merge_group`; result preserves CSV text and Excel formula.

## Task 2: Documentation, review and release
**Files:** README, RELEASE_NOTES, archived v1.3.29 notes, new spec/plan/audit.
- [x] Document narrow exception, retained safeguards and remaining unaddressed high-ratio worksheet limitation.
- [x] Compare unaffected code: classify, GUI, merge and build settings.
- [ ] Publish exact current GitHub baseline with fast-forward only.
- [ ] Verify Windows entire test suite, EXE smoke and released assets; otherwise report specific blocker.
