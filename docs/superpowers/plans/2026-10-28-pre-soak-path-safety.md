# Pre-soak Path Safety Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans, test-driven-development, and verification-before-completion task-by-task.

**Goal:** Prevent a redirected Windows ancestor from being misclassified as a fixed-local file for full content fingerprinting.

**Architecture:** A small dedicated predicate examines path components via no-follow metadata; the existing fixed-drive detector consults it only after GetDriveTypeW reports DRIVE_FIXED. On uncertain paths the existing bounded sampling remains unchanged.

**Tech Stack:** Python 3.12/3.13 standard library (pathlib, stat, ctypes), unittest, GitHub Windows Actions.

**Spec:** `docs/superpowers/specs/2026-10-28-pre-soak-path-safety-design.md`

## Global Constraints
- No user-visible behavior or workbook format changes.
- Do not resolve symlinks or probe remote data while deciding whether to use full hashing.
- Keep small file full-hash behavior, large unknown/SMB 17-window 64KiB behavior.
- Do not automatically run an 8-hour soak; it is the user's final step with Codex on authorized Windows PC.

## Review Focus
- Intermediate directory symlink; metadata-only check must detect.
- Reparse directory without symlink mode (Windows junction/cloud); flag must be honored.
- Missing/inaccessible component; fail closed into bounded sampling.
- Normal fixed-local source still hashes completely.
- 122 existing tests and Windows CI stay green.

### Task 1: Test path lineage cases
**Files:** Modify `tests/test_source_fingerprint.py`.

**Interfaces:** New `classify._has_redirecting_path_component(path: Path) -> bool`; `_is_fixed_local_drive` uses it.

- [x] Add test for a real parent symlink and a simulated reparse attribute on an ordinary parent.
- [x] Run the target tests; expected RED due to missing helper or incorrect branch.
- [x] Implement minimal metadata-only helper in `src/excel_sheet_merger/classify.py` and call from `_is_fixed_local_drive()`.
- [x] Run targeted and full unittest suite; expected PASS.
- [x] Commit changes on feature branch.

### Task 2: Windows-specific coverage and publication
**Files:** `tests/test_source_fingerprint.py`, `README.md`, `RELEASE_NOTES.md`, the resulting re-audit document.

- [x] Add Windows-only real junction integration test with safe skip if junction creation is unavailable.
- [x] Run all tests and static syntax checks; compare unaffected files by Git blob SHA.
- [x] Review for false positives and remote requests; document remaining uncertainty and put 8-hour Windows soak last.
- [ ] Publish single verified Git commit and check Windows Actions/EXE release before claiming published success.

### Task 3: Result-folder callback resilience
**Files:** `src/excel_sheet_merger/gui.py`, `tests/test_gui_reliability.py`.

- [x] Add a GUI test showing folder creation permission errors previously propagated into Tk callbacks (RED).
- [x] Catch only `OSError` around folder creation and Explorer opening; use existing Korean diagnosis and log entry.
- [x] Add Windows-only `os.startfile` failure regression without launching Explorer.
- [x] Run targeted and full GUI test suites. No UX/control-flow changes on normal use.

### Task 4: Final regression, documentation, publication
**Files:** `README.md`, `RELEASE_NOTES.md`, `docs/superpowers/stability-recheck-2026-10-28-pre-soak.md`.

- [x] Verify tests and syntax, inspect diff and unchanged merge engine SHA.
- [x] Capture remaining risk items and postpone long-run corporate Windows verification as the final Codex stage.
- [ ] Push exact tested code, confirm Windows CI 100% and EXE release assets; if blocked, report precisely without inventing success.
