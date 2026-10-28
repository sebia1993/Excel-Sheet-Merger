# Stability Disk and Long-Run Follow-up Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans, with red/green test-driven-development on every behavior change.

**Goal:** Improve low-disk failure safety and diagnostics in mixed CSV+Excel merges, while measuring long-session memory without changing feature behavior.

**Architecture:** Use the existing output-side scratch context and temporary ZIP archives. Before constructing the final ZIP, query the result drive's available bytes and compare with a documented conservative estimate; classify actual/early disk exhaustion in the existing user advice. Keep GUI logs fully scrollable and build an opt-in reproducible memory measurement instead of truncation.

**Tech Stack:** Python 3.12, openpyxl 3.1.5, ttkbootstrap, unittest/Xvfb, Windows Actions/PyInstaller.

**Spec:** `docs/superpowers/specs/2026-10-28-stability-disk-and-longrun-design.md`

## Global Constraints
- No user-visible feature, UI layout, input/output format, worksheet data or naming changes.
- No early disk guard in CSV-only path, no cross-filesystem temp moves or scratch deletion outside the owned directory.
- Only fail when a supported disk_usage query clearly suggests inadequate space; on query errors, let actual OS writes decide.
- Preserve original results, no-clobber publication and source fingerprints.

## Review Focus
- `disk_usage` low-free -> clean, actionable early failure without partial XLSX.
- `disk_usage` failure on SMB -> original output still works.
- Disk filling *after* preflight -> no published partial file and same guidance.
- A pre-existing versioned result -> byte-for-byte unchanged after failure.
- Long-run logs -> complete and in order, no silent truncation.

### Task 1: Mixed temporary archive capacity check
**Files:** `src/excel_sheet_merger/merge.py`, `tests/test_disk_capacity.py`
**Interfaces:** `_check_stitch_capacity(output_dir: Path, base: Path, csv_part: Path) -> None` called after both staging saves, before `_stitch_csv_sheets`.
- [x] Write test for low/normal/unknown free capacity, a protected previous output and no scratch residue; red failure.
- [x] Run targeted tests and note correct failure.
- [x] Implement minimal nonblocking-estimate check, use the real result drive and raise `OSError(errno.ENOSPC, ...)` only on clear low free.
- [x] Run targeted and full tests green, commit.

### Task 2: Disk failure guidance
**Files:** `src/excel_sheet_merger/failure_guidance.py`, `tests/test_failure_guidance.py`
**Interfaces:** `diagnose_failure(exc, group=None) -> FailureGuidance`, new code `DISK_FULL` for `ENOSPC`/`EDQUOT`, actionable Korean message.
- [x] Add and run failing code/type-based guidance tests (including wrapped causes).
- [x] Implement specialized case before fallback, preserve all existing guidance mappings.
- [x] Run full tests and verify no changes to existing DRM/CSV errors, commit.

### Task 3: Repeat diagnostics, documentation and publish
**Files:** `benchmarks/measure_round3_resources.py`, `README.md`, `RELEASE_NOTES.md`, `docs/superpowers/stability-recheck-2026-10-28-round3.md`.
- [x] Measure long GUI event handling and document exact memory/time with environment and limits; no log truncation.
- [x] Run full Linux/Xvfb tests, compile check and workload test, record results.
- [x] Compare Git blob hashes of untouched user-facing logic, manually review the diff and address regressions.
- [ ] Publish verified commit to `main` (user's standing approval); confirm Windows Actions full tests and actual EXE smoke, tagged release and assets before claiming Windows success.
