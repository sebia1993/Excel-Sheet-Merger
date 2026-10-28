# Readable Merge Failure Guidance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task, test-driven-development first.

**Goal:** Show failed filename, probable cause, and safe recovery action inside Windows GUI instead of only technical CP949 logs.

**Architecture:** Wrap errors at source-file boundary for precise filename and exception provenance; pure error diagnosis returns Korean guidance; worker queues failure events; GUI shows persistent banner and a detailed completion dialog while preserving full logs.

**Tech Stack:** Python 3.12, openpyxl, tkinter/ttkbootstrap, unittest, PyInstaller.

**Spec:** `docs/superpowers/specs/2026-10-28-readable-merge-failure-guidance.md`

## Global Constraints

- No changes to successful CSV/Excel parsing, group selection, time-sheet naming, or numbered output logic.
- Do not claim CP949 failures are necessarily encrypted. No circumvention instructions.
- GUI widgets accessed on Tk main thread only; Windows EXE dependency list unchanged.

## Review Focus

- Name a specific CSV within a group when CP949 decoding fails, not just the output workbook.
- Preserve original exception chain and `ValueError` compatibility with existing tests.
- Keep warning prompt readable with multiple errors, without losing details in logs.
- Do not leak errors from previous runs into a new run, and do not block independent successful groups.
- Generic/unexpected errors must not be falsely classified as DRM.

### Task 1: File-specific error provenance and guidance

**Files:** `src/excel_sheet_merger/merge.py`, `src/excel_sheet_merger/failure_guidance.py`, `tests/test_failure_guidance.py`.

**Interfaces:** `SourceFileError(ValueError)` carries `filename`, preserves `str(original)`; `diagnose_failure(error: Exception, group: MergeGroup|None) -> FailureGuidance`.

- [x] Add tests using real corrupted CSV bytes for source filename, CP949 root cause, clear non-certain DRM wording, other error categories, successful CSV regression.
- [x] Run tests and verify expected RED.
- [x] Implement minimal file context wrapper and pure mapping from cause chain to Korean detail, without touching normal CSV/Excel operations.
- [x] Run targeted tests and full suite green.
- [x] Commit.

### Task 2: UI failure banner, dialog and log

**Files:** `src/excel_sheet_merger/gui.py`, `src/excel_sheet_merger/ui_theme.py`, `tests/test_ui_design.py`, `tests/test_distribution.py`.

**Interfaces:** New `failure` event carries `FailureGuidance`; existing `done` `(completed, failed, renamed)` event unchanged, `_process_group` return tuple unchanged.

- [x] Add failing GUI tests with actual group-level failures: banner, popup, log, success of other groups, reset on next run/scan.
- [x] Run and observe RED.
- [x] Add worker guidance records/events and main-thread visible banner & Korean end-of-job messages; preserve existing interaction.
- [x] Run targeted GUI, distribution and full tests.
- [x] Commit.

### Task 3: User docs and publication

**Files:** `README.md`, `RELEASE_NOTES.md`, `docs/superpowers/...`.

- [ ] Document CP949/protection uncertainty and approved workflow, in-app guidance UX.
- [x] Execute full headless suite, compile, review diff / protected feature regression.
- [ ] Publish via GitHub plugin to main (user requested continuation on this repo) using exact tested blob hashes.
- [ ] Verify Windows Action test, EXE and release assets SHA checksum; don't claim deployment success without confirmation.
