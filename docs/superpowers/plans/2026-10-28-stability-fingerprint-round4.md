# Stability Fingerprint Round Four Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans and test-driven-development to implement this plan task-by-task.

**Goal:** Detect more metadata-preserving edits of large source files without broadening the normal user-facing behavior or doing costly whole-file SMB hashing.

**Architecture:** Preserve the existing snapshot and fingerprint API; increase large-file sample positions while keeping small-file full hashing. Reuse pre-read, post-read and pre-publish `_verify_source` calls. Add deterministic real-file regression and bounded-IO tests.

**Tech Stack:** Python 3.12+, unittest, pathlib, BLAKE2b; no new dependencies.

**Spec:** `docs/superpowers/specs/2026-10-28-stability-fingerprint-round4.md`

## Global Constraints
- CSV, XLSX, XLSM, XLS behavior and GUI unchanged.
- Keep <=8MiB full hash; above threshold sample at exactly 17 evenly spaced 64KiB positions.
- Keep read IO per check under 1.1MiB for large input.
- No new user setting or dependencies.

## Review Focus
- Unsampled fraction mutation with preserved metadata must be detected at selected sample position.
- Another quarter must be covered, not only the center.
- Small files still hash all bytes, including late changes.
- A large unmodified source must pass `_verify_source` and preserve XLSX content.
- SMB IO limited by exact read-call count and byte count.

### Task 1: Regression tests for large files
**Files:** `tests/test_source_fingerprint.py`.
**Interfaces:** `file_fingerprint(path: Path, size: int) -> bytes` and `_verify_source(item: SheetInput) -> None`.
- [x] Write failing tests for same-size same-mtime edits around 25% and 75% in a >8MiB source, and a bounded read-count test for large fingerprints.
- [x] Run targeted unittest, confirm quarter mutation test fails with current three-region algorithm.
- [x] Commit tests after red verification.

### Task 2: Bounded sampling implementation
**Files:** `src/excel_sheet_merger/classify.py`.
**Interfaces:** Unchanged signature, still BLAKE2b digest_size=20.
- [x] Sample 17 evenly spaced non-overlapping 64KiB regions, including first and final; include offsets and byte content in digest.
- [x] Run targeted + entire unittest suite with Xvfb, expect 100% pass.
- [x] Benchmark 3 vs 17 sampled windows and full-file hash on synthetic data; disclose tradeoff and remaining limitations.
- [x] Commit stable change.

### Task 3: Documentation, audit and publication
**Files:** `README.md`, `RELEASE_NOTES.md`, `docs/superpowers/stability-recheck-2026-10-28-round4.md`.
- [x] Describe stronger sampling and explicitly state it is not a full-file guarantee.
- [x] Fresh full tests, syntax checks, ZIP/XLSX example, source diff review, clean git state.
- [ ] Push tested files as one GitHub commit only if appropriate and authorized; verify Windows Actions, EXE release assets and exact SHA. Do not claim unverified completion.
