# Full Local Fingerprint Stability Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task, using test-driven-development and verification-before-completion.

**Goal:** Detect stealth modifications outside the existing 17 samples for moderately large local Windows files without adding full SMB scanning.

**Architecture:** Add conservative Windows GetDriveTypeW detection for fixed local drives; full hash only when size is 8-32 MiB on a confirmed local drive. Keep original bounded 17-window path for everything else.

**Tech Stack:** Python 3.12/3.13 standard library (pathlib, hashlib, ctypes), unittest, Windows Actions.

**Spec:** `docs/superpowers/specs/2026-10-28-local-fingerprint-completeness.md`

## Global Constraints
- No change to visible UI or file merge behavior, output data, grouping, numbering, or company DRM.
- No dependencies; full hashing must stream in <=128 KiB blocks.
- Unknown/UNC/mapped/non-fixed paths remain sampled and never trigger network probing.
- Preserve original file and result safety.

## Review Focus
- In-between-sample stealth changes with metadata restored: full-fixed hash must catch.
- SMB and mapped drives: no full hashing, no share lookup in drive classification.
- Files >32 MiB: preserve capped 17-window read.
- Unknown or inaccessible drive classification: fall back to bounded sampling.
- Existing normal merge, Excel data, errors, checkboxes unchanged.

### Task 1: Define source consistency cases
**Files:** Modify `tests/test_source_fingerprint.py`.
**Interface:** `file_fingerprint(path: Path, size: int) -> bytes`, `_is_fixed_local_drive(path: Path) -> bool`.
- [x] Add 17 MiB stealth mutation test forcing local drive; assert `_verify_source` raises a changed-file exception.
- [x] Add a bounded read test for network/unknown mode and a Windows local fixed drive characterization.
- [x] Run targeted suite and observe expected RED from unsampled mutation.

### Task 2: Conservative Windows fixed-disk selection
**Files:** Modify `src/excel_sheet_merger/classify.py`.
- [x] Add `_is_fixed_local_drive(path)` with `GetDriveTypeW` and conservative false fallback.
- [x] Full hash 8-32 MiB only when local fixed drive is confirmed, using existing 128 KiB iterator.
- [x] Run targeted suite GREEN and full tests GREEN.

### Task 3: Reporting and delivery
**Files:** Modify `README.md`, `RELEASE_NOTES.md`; create `docs/superpowers/stability-recheck-2026-10-28-round6.md`.
- [x] Record exact safe behavior, bandwidth trade-off and remaining SMB gap.
- [x] Run full tests, check syntax, local synthetic benchmark; compare release source to local hashes.
- [ ] Atomic GitHub main update after full verification, CI Windows package+EXE smoke and release asset verification; retain previous release.
