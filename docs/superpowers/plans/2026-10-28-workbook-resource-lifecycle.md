# Workbook Resource Lifecycle Implementation Plan

> **For agentic workers:** use superpowers:executing-plans and test-driven-development.

**Goal:** Release only this merge's worksheet XML and file handles on all ordinary exit paths without changing successful output.
**Architecture:** A pinned openpyxl resource adapter plus explicit source stream ownership; conversion logic remains in merge.py.
**Tech Stack:** Python 3.12, openpyxl 3.1.5, unittest, GitHub Windows Actions.
**Spec:** `docs/superpowers/specs/2026-10-28-workbook-resource-lifecycle.md`.

## Global Constraints
No input edits, user-data cleanup, feature changes, safety-limit relaxation or global library monkeypatching. Preserve output naming and primary error information. Never claim company Windows/SMB or eight-hour validation.

## Review Focus
1. A later source fails after earlier streamed rows: delete only owned worksheet temp files.
2. ZIP writing fails with ENOSPC: clean XML and close output archive without masking ENOSPC.
3. Workbook parsing fails: close source stream even if traceback remains live.
4. Another workbook is active: do not touch its temp files.
5. Successful serialization: same worksheet/style/package contents; unchanged CSV text/formula semantics.

### Task 1: test-first failure characterization
Files: `tests/test_workbook_resource_cleanup.py`.
- [x] Add real input-failure, ZIP-write-failure, parse-handle and unrelated-writer tests.
- [x] Run against exact baseline merge.py; eight assertions fail due to retained resources, not import/setup errors.

### Task 2: resource ownership fixes
Files: `src/excel_sheet_merger/workbook_resources.py`, `src/excel_sheet_merger/merge.py`.
Interfaces: `save_workbook_safely(book, path)` and `close_workbook_resources(book)`; consume current Workbook/worksheet writers, no shared state mutation outside their cleanup methods.
- [x] Use owned input file handle around load_workbook.
- [x] Retain openpyxl worksheet serialization with try/finally cleanup and output ZIP closure.
- [x] Replace only save/finalization calls; eleven new tests and available engine regressions pass.

### Task 3: review and verification
Files: spec, plan, `docs/superpowers/stability-recheck-2026-10-28-round11.md`, release notes.
- [x] Check successful package XML equality and repeated synthetic failure/success cycles.
- [x] Self-review (no independent subagent available); document remaining limitations.
- [ ] Publish only exact changed files via non-forced expected-head update.
- [ ] Verify full Windows tests, EXE smoke and release assets against actual commit; otherwise report incomplete publication truthfully.

## Local environment limitation
The container cannot resolve external DNS to install ttkbootstrap or clone the repository. Relevant source modules were reconstructed from retained files and connector reads; merge.py, temp_workspace.py, network_io.py and xlsx_safety.py hashes matched the pinned baseline. Local verification is deliberately reported as targeted engine tests, not a full 139-test baseline suite. Full current-source regression testing is delegated to the Windows workflow, which gates Release publication.
