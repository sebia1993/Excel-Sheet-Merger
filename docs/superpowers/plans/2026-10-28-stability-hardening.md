# Stability Hardening Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans, test-driven-development and verification-before-completion per task.

**Goal:** Correct P0–P3 stability audit findings without modifying normal user-visible functionality.
**Architecture:** State consistency guards; sparse Excel copy; streamed archive assembly for mixed CSV/XLSX; locked isolated scratch; atomic log persistence; bounded Tk event loop; reproducible Windows CI.
**Tech Stack:** Python 3.12, Tkinter, openpyxl, xlrd, ttkbootstrap, zipfile, unittest, GitHub Actions and PyInstaller.
**Spec:** `docs/superpowers/specs/2026-10-28-stability-hardening-design.md`

## Global constraints
- Preserve existing successful group naming, source bytes, result workbook values, formulas, date types, styles, merged cells and frozen panes.
- Keep user controls and result naming unchanged; internal smoke-test switch only for CI.
- Do not delete temp paths without marker, age and lock verification.

## Review focus
- Raced source replacement with unchanged size; source file moved after scan.
- Mixed archive with interleaved CSV/Excel and hyperlinks/formulas/merged cells.
- Other process active scratch dir older than retention.
- Log failure after successful XLSX creation.
- Main-thread event order when thousands of events queued and done is at end.

### Task 1: P0 scan state and input snapshot
**Files:** `src/excel_sheet_merger/gui.py`, `classify.py`, `merge.py`, `tests/test_stability.py`.
**Interfaces:** Preserve `scan_folder`, `MergeApp._scan`, `merge_group` API; extend `SheetInput` with optional snapshot default.
- [x] Write tests: A success/B scan failure and changed file snapshots, observe red.
- [x] Fix invalidation and scan/output binding; snapshot verification; run tests green.
- [x] Commit.

### Task 2: P1 memory and sparse Excel
**Files:** `src/excel_sheet_merger/merge.py`, `tests/test_stability.py`.
**Interfaces:** `_append_openpyxl` sparse cell enumeration; mixed CSV streaming/stitching; existing `merge_group` signature.
- [x] Write tests for sparse grid, mixed files with source formatting, observe red.
- [x] Implement sparse cell copy and streamed worksheet insertion with bounded memory.
- [x] Run full suite and representative performance test green; commit.

### Task 3: P1 scratch recovery, P2 log, UI event loop
**Files:** `temp_workspace.py`, `merge.py`, `gui.py`, `tests/test_stability.py`.
- [x] Write tests for orphan recovery, active locks, failed log save and queue batching, observe red.
- [x] Implement owner-locked scratch cleanup, atomic log and bounded event draining.
- [x] Run full suite green; commit.

### Task 4: P2-P3 Windows release resilience
**Files:** `app.py`, `requirements-build.txt`, `.github/workflows/windows-exe.yml`, `tests/test_stability.py`, docs.
- [x] Add executable smoke path and tests, pin dependencies/Windows runner, collect frozen requirements and SHA256.
- [x] Run full tests/py_compile; self review; push branch and GitHub Windows workflow; confirm release assets (pending GitHub publication).
- [x] Final independent re-audit, document unresolved risks.
