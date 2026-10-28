# Unique Result File Naming Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans with test-driven-development for each task.

**Goal:** Preserve every existing XLSX result and create an extra suffix-numbered XLSX when the name exists, with clear GUI notices.

**Architecture:** Keep the workbook generated in a unique temporary XLSX; publish atomically under the first unused filename. The merge engine returns status and real path for the GUI to log without touching file classification or workbook data.

**Tech Stack:** Python 3.12+, openpyxl, unittest, ttkbootstrap/Tkinter, GitHub Actions Windows/PyInstaller.

**Spec:** `docs/superpowers/specs/2026-10-28-duplicate-output-versioning-design.md`

## Global Constraints
- Preserve existing CSV/excel reading, checkbox selection, and timestamp sheet names.
- Never overwrite source or results; keep atomic cleanup.
- EXE must remain standalone, no extra runtime dependencies.

## Review Focus
- Existing XLSX appears between validation and publication -> atomic retry.
- Existing renamed version `(2)` -> suffix increments to `(3)`.
- Broken source when a result exists -> error, no new output.
- Both CSV-only and mixed CSV/Excel -> identical semantics.
- GUI shows renamed success, not skipped; original filename visible in record.

### Task 1: Engine collision handling
**Files:** `src/excel_sheet_merger/merge.py`, `tests/test_merge.py`, `tests/test_excel_inputs.py`.

**Interfaces:** `merge_group(group, output_dir) -> MergeResult`, `.path` actual output, `.status` `CREATED|RENAMED`.

- [x] Write tests for second/third output suffixes, preservation, error cleanup, mixed input, atomic collision. Fails on `MergeStatus.RENAMED` being unavailable or skipped output.
- [x] Run targeted tests and observe failure.
- [x] Replace skip-on-exists with a single atomic loop (`os.rename` on Windows; `os.link` on POSIX) and status result.
- [x] Run tests and verify workbook contents and protected output unchanged.
- [x] Commit task.

### Task 2: UI notices and progress
**Files:** `src/excel_sheet_merger/gui.py`, `tests/test_distribution.py`, `tests/test_ui_design.py`.

**Interfaces:** Keep `_process_group(..., completed, failed, skipped, records) -> (int,int,int)` and `done` queue tuple shape; repurpose third counter as `renamed` and update text.

- [x] Add failing test for rename notifications, complete count, log text and returned renamed counter.
- [x] Run targeted tests to observe failure.
- [x] Make `_process_group` log old/new names and `_consume_events` display renamed count.
- [x] Run tests for GUI selection and worker events.
- [x] Commit task.

### Task 3: Documentation, release checks
**Files:** `README.md`, `RELEASE_NOTES.md`; this design + plan docs.

- [x] Update user documentation with three example filenames, atomic safety and UI wording.
- [x] Run full `xvfb-run -a python -m unittest discover -s tests -v` and Python compile checks.
- [x] Compare `classify.py`/theme hashes before and after.
- [ ] Push exact tested source to GitHub and verify remote commits/blobs.
- [ ] Confirm Windows GitHub Actions succeeds; verify published EXE and release assets, otherwise report exact blocker.
