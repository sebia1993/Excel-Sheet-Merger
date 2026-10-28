# Incomplete Source Read Stability Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans with test-driven-development and verification-before-completion.

**Goal:** Reject incomplete source fingerprint reads instead of implicitly trusting a short stream.

**Architecture:** Add byte-count validation to full-file hashes, exact 64-KiB-window collection for sampled hashes; preserve digest inputs, offsets and existing classification error behavior. Cover both boundaries with tests and a complete GUI/merge regression suite.

**Tech Stack:** Python 3.12 and Python 3.13, standard `hashlib` / buffered file IO, unittest, GitHub Actions Windows.

**Spec:** `docs/superpowers/specs/2026-10-28-fingerprint-short-read-safety.md`

## Global Constraints
- Do not change normal output contents, UI, group checkbox, duplicate suffix logic, security policies or supported formats.
- Keep sampled read budget 17 × 64KiB; full hash bounded per 128KiB read.
- Preserve PermissionError fallback from company DLP/DRM in scanner; do not bypass access control.

## Review Focus
- Full stream ends prematurely before recorded size: fail closed.
- Full stream grows beyond recorded size: fail closed.
- Sample window is missing bytes: fail closed even with metadata unchanged.
- Short but nonempty sample reads that complete on a subsequent read: produce same fingerprint as one full read.
- Existing normal CSV and Excel, network error explanation and atomic output remain unchanged.

### Task 1 — Red/green file fingerprint consistency
**Files:** `tests/test_source_fingerprint.py`, `src/excel_sheet_merger/classify.py`.
**Interfaces:** `file_fingerprint(path: Path, size: int) -> bytes` stays unchanged.
- [x] Write synthetic failing tests for premature EOF in both modes and partial-but-complete sample reads.
- [x] Run them before implementation, record failures due to existing behavior.
- [x] Add exact expected byte-count verification and sampled read-completion loop.
- [x] Run targeted tests and full regression suite; inspect output.
- [x] Commit.

### Task 2 — Document and publish only after verification
**Files:** `README.md`, `RELEASE_NOTES.md`, `docs/superpowers/stability-recheck-2026-10-28-round7.md`.
- [x] Document truncated-read handling and residual limitations.
- [x] Run all tests with Xvfb and Python compile; verify normal mixed XLSX output.
- [ ] Compare tested file blobs to remote GitHub objects, publish with no-force lease.
- [ ] Confirm Windows tests, packaged EXE smoke, Release download, hash and old release retention.
- [ ] Conduct a second audit of residual risks, report what remains.
