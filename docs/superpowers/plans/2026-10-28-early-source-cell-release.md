# Excel Source Cell Lifetime Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans and test-driven-development.

**Goal:** Release in-memory source cells during Excel-to-Excel copying to reduce peak RSS while preserving workbook output.

**Architecture:** In `merge._append_openpyxl`, consume a loaded source worksheet's `_cells` dictionary incrementally. Copy each cell to the output using the same code paths as before; retain independent worksheet metadata references and close the loaded source workbook in its existing `finally` block.

**Tech Stack:** Python 3.12, openpyxl 3.1.5, unittest, PyInstaller Windows Actions.

**Spec:** `docs/superpowers/specs/2026-10-28-early-source-cell-release-design.md`

## Global Constraints
- No visible GUI/function/file-naming change; preserve all XLSX values, formulas, styles and merged geometry.
- No new production dependencies, no input-file modifications.
- Retain ZIP archive safety checks and source-fingerprint validation.

## Review Focus
- Sparse cells, merged cells, styled empty cells and borders must still copy.
- Source comments/hyperlinks and source worksheet independent metadata must survive.
- Source validation or copy exceptions must close the source workbook.
- Formula-looking CSV still text, Excel formula still formula.
- Source workbook on disk must stay byte-for-byte unchanged.

### Task 1: Characterize memory lifetime before modifying production
**Files:** Create `tests/test_memory_lifecycle.py`; modify `src/excel_sheet_merger/merge.py`.

- [x] Write a failing test that spies on actual copy calls and verifies the source cell dictionary shrinks before the transfer reaches its midpoint.
- [x] Run the test and observe assertion fail with the unchanged full-sized dictionary.
- [x] Replace non-mutating iteration with incremental entry consumption of the private loaded source workbook, retaining the exact cell copy body.
- [x] Run targeted and entire unittest suite with Xvfb; ensure unchanged workbook semantics.
- [x] Capture before/after subprocess peak RSS for the same synthetic workbook files.
- [ ] Commit code and tests.

### Task 2: Documentation, re-audit, and packaging
**Files:** `README.md`, `RELEASE_NOTES.md`, `docs/superpowers/stability-recheck-2026-10-28-round10.md`.

- [x] Document measured memory trade-offs and exact limitations (Windows/SMB/DRM, 8-hour soak not performed).
- [x] Confirm only source lifetime changed; existing `classify.py`, `gui.py`, `xlsx_safety.py` stay hash-identical.
- [ ] Commit docs. Publish tested files to GitHub only via non-forced fast-forward verified against upstream SHA.
- [ ] Check all Windows Actions steps including packaged EXE smoke and new release assets. Do not claim CI success without evidence.
