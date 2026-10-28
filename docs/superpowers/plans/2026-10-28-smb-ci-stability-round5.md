# Network Error and Node24 CI Stability Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement task-by-task, with superpowers:test-driven-development.

**Goal:** Keep existing merger output safe under explicit network disconnections and update CI first-party actions to Node24.

**Architecture:** Centralize OS network error recognition in a pure helper, use it in classification and diagnostic layers, and preserve fail-closed publishing. Update the CI action version references only, with the same Windows runner and Python packaging tests.

**Tech Stack:** Python 3.12, unittest, Tkinter, openpyxl 3.1.5, GitHub Actions Windows 2025.

**Spec:** `docs/superpowers/specs/2026-10-28-smb-ci-stability-round5.md`

## Global Constraints
- No new features or visual changes; no source data modification or output overwriting.
- No background network retries, especially after ambiguous publishing.
- No new third-party runtime dependencies; keep Python 3.12.10 and all pinned wheel versions unchanged.
- Preserve `MergeStatus` semantics, group checks, sheet naming, and existing Korean failure dialog layout.

## Review Focus
- DLP permission without network outage must not become an excluded group.
- Network loss halfway through scanning should not leave a partial group visible.
- Network loss while writing must not leave partial named XLSX or delete prior result.
- Possible publication-before-timeout requires a caution to inspect result before retry.
- CI version changes must retain full suite, EXE packaging, and artifact/release verification.

### Task 1: Network failure recognition and safe scan
**Files:** `src/excel_sheet_merger/network_io.py`, `src/excel_sheet_merger/classify.py`, `tests/test_network_io.py`.

- [x] Add failing tests for errno + winerror classification, non-network exceptions, early failure on disconnected fingerprint, and continued treatment of PermissionError.
- [x] Run targeted tests and confirm RED for required behavior.
- [x] Implement `is_network_io_error(error: BaseException) -> bool` and re-raise explicit network errors in `scan_folder` instead of discarding the fingerprint.
- [x] Run targeted tests, then full suite.
- [x] Commit reviewed task.

### Task 2: Network-aware guidance and atomicity regression
**Files:** `src/excel_sheet_merger/failure_guidance.py`, `src/excel_sheet_merger/gui.py`, `tests/test_network_io.py`.

- [x] Add failing tests for source failure, write failure, ambiguous publish and non-network errors; assert no automatic retry and no existing result overwrite.
- [x] Run targeted tests, observe RED.
- [x] Add `NETWORK_IO` failure guidance above CSV/Excel generic parsing classifications and reuse it in scan-failure popup; preserve existing error categories.
- [x] Run targeted and full suite; inspect code diff for unchanged success path.
- [x] Commit reviewed task.

### Task 3: Node 24 CI migration, regression & docs
**Files:** `.github/workflows/windows-exe.yml`, `tests/test_distribution.py`, `README.md`, `RELEASE_NOTES.md`, `docs/superpowers/stability-recheck-2026-10-28-round5.md`.

- [x] Add a failing workflow regression assertion for checkout v5, setup-python v6, upload-artifact v6 and preserved Windows exe smoke.
- [x] Run targeted test, confirm RED.
- [x] Update only the three action majors, document risks and test scope.
- [x] Run full suite + Python compilation + approved synthetic fault injections.
- [ ] Publish exact tested blobs in one commit to `main`, validate GitHub Actions Windows run, EXE packaged smoke, new release assets and SHA; if unavailable/failing report blocker instead of claiming success.
