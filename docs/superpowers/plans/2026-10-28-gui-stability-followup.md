# GUI Stability Follow-up Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans with test-driven-development to implement this plan task-by-task.

**Goal:** Recover UI state on worker thread startup failure and batch queued GUI log rendering without altering merge output.

**Architecture:** Keep existing GUI and worker event protocol; protect only worker startup transition and replace repeated consecutive `Text.insert` calls with a single batch insertion. All workbook modules remain untouched.

**Tech Stack:** Python 3.12, Tkinter/ttkbootstrap, unittest, GitHub Actions Windows runner.

**Spec:** `docs/superpowers/specs/2026-10-28-gui-stability-followup.md`

## Global Constraints
- No parser, file naming, CSV/Excel output, visual layout or existing displayed log history changes.
- No auto-retry of failed thread start.
- All existing tests pass and new tests fail against the baseline before fixing.

## Review Focus
- Resource exhausted thread start leaves busy True: regression must prove UI recovered.
- Initial folder/group selection intact after failed thread start.
- Consecutive log burst preserves all lines and formatting, including newest scrolled into view.
- Log lines before and after progress/failure/done remain ordered.
- Windows EXE packaging still opens and merges correctly in CI; no new dependency.

### Task 1: Thread startup error recovery

**Files:** `src/excel_sheet_merger/gui.py`, `tests/test_gui_reliability.py`.
**Interfaces:** existing `MergeApp._start_merge(selection_only: bool) -> None`, `MergeApp.busy` and button states; `messagebox.showerror` for UI guidance.

- [ ] Add failing GUI test simulating a thread `start` RuntimeError and asserting `busy=False`, original buttons reenabled, error popup and subsequent normal start possible.
- [ ] Run targeted test: expected failure on baseline.
- [ ] Implement narrow exception recovery around Thread construction and start.
- [ ] Run targeted test: expected pass; run all tests.

### Task 2: Text log rendering batch

**Files:** `src/excel_sheet_merger/gui.py`, `tests/test_gui_reliability.py`.
**Interfaces:** existing `_write_log`, `_consume_events`, event queue. Shared formatting `[%H:%M:%S] message\n`.

- [ ] Add failing GUI test enqueuing 100 logs and verifying exactly one Text insert plus all 100 visible in order.
- [ ] Add test for log-progress-done ordering, error notices and final UI state.
- [ ] Run targeted test: expected failure on baseline.
- [ ] Implement `_append_log_lines` and buffered flush on event boundaries; preserve existing public `_write_log`.
- [ ] Run targeted tests and full suite; verify `git diff --check` and py_compile.

### Task 3: Distribution and second review

**Files:** `README.md`, `RELEASE_NOTES.md`, existing Windows Actions workflow unchanged, docs.

- [ ] Record scope/limits; verify source modules `classify.py`/`merge.py` unchanged by Git blob hash.
- [ ] Run 30+ repeated CSV/XLSX sample merges (bounded soak, no false 8-hour claim).
- [ ] Review code for additional risks separately.
- [ ] Upload same tested files to GitHub, verify blob SHA and Windows Actions; only then report new release URL.
