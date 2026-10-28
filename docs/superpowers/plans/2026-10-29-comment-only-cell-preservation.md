# Blank Cell Comment Preservation Implementation Plan

> For agentic workers: Use superpowers:executing-plans and test-driven-development.

**Goal:** Prevent silent loss of comments attached to unstyled blank Excel source cells.

**Architecture:** Only change the early-skip predicate in _append_openpyxl. Keep every other input/output/GUI behavior.

**Tech Stack:** Python 3.12, openpyxl 3.1.5, unittest, Windows PyInstaller.

**Spec:** docs/superpowers/specs/2026-10-29-comment-only-cell-preservation.md

## Global constraints
Existing CSV/XLSX/XLSM/XLS handling, checked groups, numbered outputs, all logs and original file protection remain identical. Eight-hour Windows soak postponed.

## Review focus
1. A blank unstyled cell with a note is copied without giving it a data value.
2. Ordinary empty cells are still skipped, avoiding sparse-grid explosion.
3. Workbook with no actual data still fails the previous empty-sheet rule.
4. Formula/style/merge comments and links on populated cells are preserved.
5. Original XLSX bytes and already published results remain untouched.

## Task 1: Red/green regression
**Files:** tests/test_memory_lifecycle.py, src/excel_sheet_merger/merge.py.
**Interface:** _append_openpyxl(ws, path: Path) -> None.
- [x] Create direct real-workbook note-only blank-cell test and observe old code failure: comment unexpectedly None.
- [x] Replace skip predicate to test absence of value, style, comment AND hyperlink.
- [x] Verify direct test and real CSV+XLSX group merge including second output (2), source byte preservation.
- [x] Run full Linux/Xvfb unittest 157 with 4 Windows-only skips; syntax and diff verification.

## Task 2: Documentation/release
**Files:** README.md, RELEASE_NOTES.md, docs/releases/v1.3.28.md, new spec/plan/re-audit.
- [x] Document reproduced defect and scoped repair, log-memory benchmark, unchanged UI, and deferred 8h soak.
- [ ] Create one Git data commit against verified remote main and ensure tested source and tests match.
- [ ] Check full Windows Actions tests, packaged EXE smoke and release attachments before claiming final completion.
