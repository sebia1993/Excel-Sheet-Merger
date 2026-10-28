# Folder Path Resolution Failure Recovery Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans and test-driven-development, in the existing isolated worktree.

**Goal:** Avoid Tk callback errors and stale preview publication when `Path.resolve()` fails after scanning or before starting a merge.

**Architecture:** Make the UI state commit contingent on both scan and path resolution; on merge start error invalidate prior state, display existing failure warning, avoid starting worker. No merge-engine changes.

**Tech Stack:** Python 3.12, tkinter, ttkbootstrap, unittest/Xvfb, Windows GitHub Actions/PyInstaller.

**Spec:** `docs/superpowers/specs/2026-10-29-path-resolution-io-safety.md`.

## Global constraints

No normal-behavior, UI layout, or conversion changes. Path failures must never enable a stale merge. Windows 8-hour soak remains deferred.

## Review focus

1. Scan succeeds but path resolution fails: no half-state.
2. Existing preview then merge start resolution fails: invalidated preview and disabled execution.
3. Windows network disconnect error has actionable Korean message.
4. Symbolic link loop (RuntimeError) cannot strand UI in previous state.
5. Normal scan/start still works; no new output from failing operations.

### Task 1: Add red-green path-resolution regressions

**Files:** `tests/test_path_resolution_safety.py`, `src/excel_sheet_merger/gui.py`.

**Interfaces:** `_scan() -> None`, `_start_merge(selection_only: bool) -> None`, `_invalidate_scan() -> None`.

- [x] Add 3 Tk-based tests simulating `Path.resolve()` OSError at scan completion, OSError and RuntimeError at merge-start. Assert no stale groups/worker and recovery text.
- [x] Run targeted tests before production edits and verify failures are because exceptions escape.
- [x] Place scan resolution within guarded region; commit `report` only once all checks succeed.
- [x] Catch `OSError`/`RuntimeError` around start path resolution; invalidate and show error, no worker started.
- [x] Run targeted red-green, complete suite, syntax and isolated git diff checks.
- [ ] Commit and review only tested files.

### Task 2: GitHub Windows validation

**Files:** `README.md`, `RELEASE_NOTES.md`, this plan/spec/audit.

- [ ] Document changed failure semantics, preserved normal behavior, limitations.
- [ ] Upload only changed Git blob objects to current GitHub main through a lease-checked fast-forward.
- [ ] Verify full Windows Actions tests, packaged EXE smoke, uploaded Release EXE/SHA and exact head commit match.
- [ ] Report remaining risks without claiming company 8h validation.
