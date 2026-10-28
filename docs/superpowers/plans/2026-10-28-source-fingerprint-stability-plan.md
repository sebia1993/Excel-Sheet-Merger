# Source Fingerprint Stability Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans or superpowers:subagent-driven-development. Steps are test-first and preserve existing interfaces.

**Goal:** Detect source byte changes that preserve size and timestamp, with bounded I/O for large files.

**Architecture:** Track a content fingerprint in scanned `SheetInput` alongside existing OS metadata. Read a bounded or full fingerprint before/after each merged source and before publishing. Keep CSV/Excel parsers unchanged.

**Tech Stack:** Python 3.12, hashlib, unittest, openpyxl, Windows GitHub Actions.

**Spec:** `docs/superpowers/specs/2026-10-28-source-fingerprint-stability-design.md`

## Global Constraints

- No changes to filenames, sheet names, CSV text storage, existing formats or GUI layout.
- No unbounded source reading solely for fingerprinting on >8 MiB files.
- Fail closed when scanned bytes have demonstrably changed, maintain DRM-readable error messages.

## Review Focus

- Identical size + restored mtime edits must be detected for small files.
- 9 MiB sampled middle edits should be detected.
- Unchanged legitimate mixed files still publish identical contents.
- Fingerprint read permissions must not cause unsupported claims of DRM bypass.
- Repeat checks must not impose full-scan I/O on large sources.

### Task 1: Regression tests (RED)

**Files:** Create `tests/test_source_fingerprint.py`.

**Interfaces:** Consume `scan_folder`, `merge_group`, `_verify_source`, `file_fingerprint`.

- [ ] Add tests for identical-metadata mutation in a small CSV and >8 MiB sampled region; prove they fail on baseline.
- [ ] Add unchanged CSV/mixed input and fingerprint read-error fallback tests.

### Task 2: Fingerprint implementation (GREEN)

**Files:** Modify `src/excel_sheet_merger/classify.py`, `src/excel_sheet_merger/merge.py`.

**Interfaces:** Preserve `file_snapshot(path) -> 4-tuple`; produce `file_fingerprint(path, size) -> bytes`; add `SheetInput.content_fingerprint: bytes | None`.

- [ ] Implement bounded hash and consistent scan metadata handling.
- [ ] Extend `_verify_source` to compare bytes and revalidate metadata.
- [ ] Run specific tests, then entire suite and compile checks.

### Task 3: Stress and distribution

**Files:** Docs and release notes; no GUI or Excel reader changes.

- [ ] Benchmark fingerprint read time for representative small and >=8 MiB files, document tradeoff.
- [ ] Run 30+ repeated CSV/XLSX merges and validate contents/temporary cleanup.
- [ ] Self-review with independent pass and track remaining risks.
- [ ] Upload verified Git blobs to GitHub, publish Release via Windows CI on main; confirm build + EXE self-test + attached asset.
