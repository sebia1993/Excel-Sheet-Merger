# Group Checkbox Selection Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox syntax.

**Goal:** Allow explicit multi-group checkbox selection without changing Excel/CSV merging.
**Architecture:** Parent Treeview rows display check glyphs; a separate checked ID set is the only selected-merge source. Mouse and keyboard toggle groups; toolbar actions select/clear all; controls lock during work.
**Tech Stack:** Python 3.12, Tkinter/ttkbootstrap, unittest, GitHub Actions/PyInstaller.
**Spec:** `docs/superpowers/specs/2026-10-28-group-checkbox-selection-design.md`

## Global constraints
- Do not modify `classify.py` or `merge.py`.
- Initial/re-scan state: no groups checked.
- Only group parent rows can be checked; child and excluded rows are not selectable for merging.
- Existing `전체 그룹 병합` ignores checked state.
- Windows build from main automatically publishes a new Release with the existing workflow; do not alter release workflow unnecessarily.

## Review focus
- Clicking a tree expander must not toggle a group.
- Keyboard Space on a child/empty area must not add a group.
- Group checking must remain separate from Treeview selection highlighting.
- Re-analysis must clear prior state, even if IDs g0/g1 are reused.
- Toggle attempts during in-flight conversion must not change the worker's group list.

### Task 1 — checkbox state and UI
Files: `src/excel_sheet_merger/gui.py`, `tests/test_ui_design.py`.
Interfaces: `_selected_groups() -> list[MergeGroup]`, `_toggle_group(group_id: str)`, `_set_all_groups(checked: bool)`, `_on_tree_click(event)`, `_on_tree_space(event)`, `_refresh_check_controls()`.
- [x] Write GUI tests of initial state, 2 checked groups, all/none, subitems ignored, button disabled, re-scan, busy guard.
- [x] Run failing GUI suite with `xvfb-run -a python -m unittest discover -s tests -p test_ui_design.py -v`; expect new tests fail on missing state/behavior.
- [x] Implement checkbox UI and state, keeping worker code unchanged.
- [x] Run GUI suite, expect all passing.
- [x] Commit task with tests and GUI changes.

### Task 2 — event fidelity, integration and docs
Files: `tests/test_ui_design.py`, `README.md`, `RELEASE_NOTES.md`.
Interfaces: same methods as Task 1.
- [x] Write real Tk mouse/keyboard event tests, and real workbook integration test for checked-only vs all.
- [x] Run failing tests then implement any missing input behavior.
- [x] Update README / release notes for checkbox semantics; remove outdated child-selection instructions.
- [x] Run `xvfb-run -a python -m unittest discover -s tests -v`, expect all passing.
- [x] Compile, inspect diff and commit task.

### Task 3 — publish and validate
Files: no production changes required by default.
- [ ] Verify GitHub main has not moved and compare source changes.
- [ ] Publish only verified changed text files to GitHub main via connector; keep old Releases intact.
- [ ] Observe Windows workflow test/build result and confirm newly-created GitHub Release asset (EXE, SHA256SUMS).
