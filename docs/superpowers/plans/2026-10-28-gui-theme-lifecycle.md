# GUI theme lifecycle implementation plan

> Use superpowers:executing-plans and test-driven-development.

**Goal:** Fix empty scrollbar colors and stale theme ownership without changing UI/features.
**Architecture:** Existing ttk styles; per-root owned Bootstrap teardown; subprocess GUI regressions.
**Tech Stack:** Python 3.12, ttkbootstrap 2.2.6, Tk, unittest.
**Spec:** docs/superpowers/specs/2026-10-28-gui-theme-lifecycle.md

## Global constraints
No merge/classification/output-policy changes. No error suppression. No 8-hour soak.

## Review focus
Fresh process versus sequential roots; teardown ownership; fallback theme; real scroll
rendering; all existing minimum-window and error-banner assertions.

### Task 1: Regression first
Files: tests/test_gui_theme_lifecycle.py.
- [x] Add fresh-process rendering and repeated-root interpreter ownership tests.
- [x] Run targeted tests; confirm actual background error/stale interpreter failures.

### Task 2: Fix production
Files: src/excel_sheet_merger/ui_theme.py.
- [x] Set Vertical.TScrollbar focuscolor using existing primary color.
- [x] Add release_theme(root) with strict singleton/root ownership check; release
      the library named-font cache with that owned instance; bind only to the root Destroy event.
- [x] Run targeted and full suite; investigate any remaining failure, do not weaken tests.

### Task 2b: Correct exposed minimum-window layout
Files: src/excel_sheet_merger/gui.py, tests/test_gui_theme_lifecycle.py.
- [x] Keep existing test_cp949_failure_shows_filename_reason_and_action_in_gui_and_popup
      unchanged; it fails at 60px on Windows and 14px on Linux after root ownership repair.
- [x] Add a real rendered error-layout test checking minimum preview and restoration
      after resize and failure reset, with every message preserved.
- [x] Track original card padding; compact whitespace only with failure + height <700.
- [x] Run full suite and zero-background-error diagnostics.

### Task 3: Verify, review, release
Files: RELEASE_NOTES.md and docs/superpowers/stability-recheck-2026-10-28-round12.md.
- [x] Run real-theme Linux suite and isolated Windows diagnostic without publishing.
- [x] Self-review exact diff and original module hashes; independent reviewer unavailable.
- [ ] Publish only tested files onto unchanged main; verify Windows/EXE/release SHA.
      Final publication verification is recorded by the publishing commit Actions, not inferred from branch diagnostics.
