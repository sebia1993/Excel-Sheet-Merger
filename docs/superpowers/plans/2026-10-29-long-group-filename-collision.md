# Long Site Output Filename Collision Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL superpowers:executing-plans + test-driven-development.

**Goal:** Avoid artificial result-name collisions caused by the 150-character safety cap while preserving existing names for normal sites.
**Architecture:** A narrow output-name helper chooses the current `_safe_filename` result when it still has the date suffix, and a capped deterministic site+digest+date only for overflow. Existing collision rejection remains.
**Tech Stack:** Python 3.12, unittest, openpyxl 3.1.5, Windows GitHub Actions.
**Spec:** `docs/superpowers/specs/2026-10-29-long-group-filename-collision.md`.

## Constraints

Preserve normal outputs and all input conversions; no GUI changes. No 8-hour Windows soak until last.

## Review focus

- Same long site on different dates is never assigned one output name.
- Two different long sites sharing the first 150 characters remain distinct.
- Output names are <=150 characters before `.xlsx` and include the full YYYYMMDD date.
- Distinct groups retain their own source data; old results never overwritten.
- Normal sites retain exactly the same output naming.

### Task 1: TDD output naming

**Files:** `src/excel_sheet_merger/classify.py`, `tests/test_classify.py`.

- [x] Add two failing tests for long site/different dates and distinct sites sharing the prefix; test public `scan_folder` and actual `merge_group` files.
- [x] Run targeted unittest and observe original output-name collision ValueError.
- [x] Implement `_safe_group_output_name(site, date)` used only when assigning group output_name; use deterministic tail and digest on overflow.
- [x] Run targeted and full suite, Python compile check, git diff check.
- [ ] Review change and commit.

### Task 2: Documentation and Windows distribution

**Files:** README.md, RELEASE_NOTES.md, spec, plan, audit, prior release notes archive.

- [x] Document long-site behavior and reason for special suffix.
- [ ] Push tested code via lease-checked GitHub main update.
- [ ] Verify Windows 100% tests and packaged EXE repeated CSV+Excel smoke before claiming deployment.
- [ ] Re-audit remaining candidates; defer company PC eight-hour soak.
