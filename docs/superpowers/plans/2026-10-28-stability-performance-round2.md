# Stability Performance Round Two Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task, with test-driven-development.

**Goal:** Reduce the mixed CSV+Excel ZIP recompression overhead while retaining all observable program behavior, with repeat-workload guardrails.

**Architecture:** Keep normal Excel in-memory workbook and streamed CSV temp workbook; preserve the existing ZIP member copy, only reduce compression effort for replaced CSV worksheet XML. Add high-value integration tests and benchmark script.

**Tech Stack:** Python 3.12, openpyxl 3.1.5, ZIP format, unittest, GitHub Actions Windows 2025.

**Spec:** `docs/superpowers/specs/2026-10-28-stability-performance-round2.md`

## Global Constraints
- No GUI/user feature or file-format changes; existing exporter remains identical apart from compression level for CSV worksheet XML.
- Preserve source ZIP member metadata; leave normal Excel member compression as before.
- No new Python dependencies; Python/Excel installation not required on client.

## Review Focus
- Larger output may increase disk use: benchmark and report tradeoff.
- Recompressed XML must not mutate source `ZipInfo` nor invalid CRC / package relationships.
- Worksheet formulas and raw CSV text must stay typed as before.
- Several successive merges must not leak scratch files or corrupt numbered results.
- Linux tests are not substitutes for corporate Windows, network share, or 8-hour workload.

### Task 1: Performance and compatibility acceptance
**Files:** `tests/test_zip_stitch_integrity.py`, `benchmarks/measure_stitch.py`.

- [x] Create a deterministic mixed XLSX/CSV fixture verifying ZIP CRC, metadata of normal package members, output value types and styles.
- [x] Add an opt-in performance benchmark comparing existing output to equivalent level-6 zip writer and output size in synthetic sample; observe baseline slower than goal (RED).
- [x] Run targeted tests, capture output; commit test-first change.

### Task 2: Compression effort reduction
**Files:** `src/excel_sheet_merger/merge.py`.

- [x] Restrict lower compression level to the streamed CSV worksheet members using copied ZipInfo; avoid changing other members.
- [x] Run opt-in benchmark, verify speed and size; run complete test suite and syntax check; commit.

### Task 3: Repeated-output quality and packaging
**Files:** `tests/test_long_run.py`, `README.md`, `RELEASE_NOTES.md`, `docs/superpowers/stability-recheck-2026-10-28-round2.md`.

- [x] Add 100-repetition synthetic mixed workbook test without input changes; verify all worksheet names/types, all output files and no scratch residue.
- [x] Re-run full suite and bench; review tested diff; report any outstanding gaps.
- [ ] Push tested Git objects to main (user approved), validate Windows Actions full suite & packaged EXE release to a new tag; no claims if failed.
