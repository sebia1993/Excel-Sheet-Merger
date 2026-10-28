# Excel Sheet Merger — Stability Re-audit Round 10

Date: 2026-10-28. Baseline GitHub main `d6db1a03eccb7d42585559f1df024310562ff2af`, Release v1.3.25. The local source `merge.py` and `xlsx_safety.py` blob hashes matched remote. User requested automatic scoped stability improvements, keeping 8-hour Windows soak as the last stage.

## Reproduced fault

Normal-mode openpyxl instantiates every source cell before copying it into a second workbook. Previous `_append_openpyxl` iterated `source._cells.values()` without releasing any source references until the source workbook function returned, so densely populated XLSX workbooks retained both source and destination cells at peak memory. A new live-cell lifetime test observed 4,000 source cells still resident at the halfway point and failed before the fix.

## Change

- `_append_openpyxl` now consumes the private, in-memory source worksheet dictionary incrementally (`while source._cells: ... popitem()`); each source Cell is copied using *the same existing assignment/style/comment/hyperlink code* and then ceases to be held by the source worksheet dictionary. Worksheet metadata, merge coordinates and styles are copied afterwards as before. Original file on disk is unmodified.
- No changes to the CSV path, ZIP archive resource guard, grouping, filenames, result publishing, UI, error messages, network validation or corporate data policy.

## Evidence

- TDD RED: `test_live_source_cell_ownership_decreases_during_copy` failed on baseline (`4000 not less than or equal to 2500`). GREEN: new test and existing end-to-end tests passed.
- Local `xvfb-run -a python -m unittest discover -s tests -q`: 139 tests, OK (4 Windows-only skipped). Source with formula, font, comment, hyperlink, merged region, sparse far cell, styled empty cell, row and column settings and freeze pane produced correct result; source byte content remained identical.
- Benchmark: two XLSX sources each containing 12,000 rows × 8 columns (~415KiB compressed per source), combined output ~826KiB. Three *fresh subprocesses per variant*, alternating baseline and optimized, CPython 3.13/openpyxl 3.1.5 on Linux, `/usr/bin/time` resource peak RSS:
  - Baseline: 209,320 / 209,508 / 209,332 KiB; median 209,332 KiB; duration median 5.15 s.
  - Optimized: 192,728 / 192,720 / 192,768 KiB; median 192,728 KiB; duration median 4.58 s.
  - Median RSS decrease 16,604 KiB (~7.9%). Duration varies with filesystem caching; do not generalize to production Windows/SMB performance.
- ZIP content comparison of baseline and optimized outputs: identical workbook sheets, styles and worksheet XML; only `docProps/core.xml` (creation/modification timestamp) differed.

## Follow-up candidates and risks

1. P1: Actual company Windows 11/SMB/DRM/AV 8-hour stress and real-device Codex work remains **last**; no claim of completion.
2. P2: Valid but extremely repetitive Excel input may trigger existing ZIP expansion-ratio safety guard. In a synthetic 2,400-cell valid workbook with identical 30,000-character ASCII values, the worksheet XML advertised ~72MiB, ~775× compression and was blocked. This was **previously documented and intentional** to avoid decompression/memory bombs; any adjustment requires strict streaming structural validation and memory bounds, not simply relaxing the limit.
3. P2: Very complex Excel native objects (charts, pivot tables, external links, formula caches) are beyond existing fidelity guarantees; user-permitted native Excel verification remains necessary.
4. P2: Network and files >32MiB use bounded source fingerprint samples; unsampled edits may escape detection. SMB whole-file scans have measurable I/O costs.
5. P3: Tkinter visible logs grow indefinitely in a long session; truncation would alter user-visible history, so implement only after a non-lossy virtualization design and UI tests.
6. P3: Private openpyxl `_cells` mapping is pinned and used elsewhere in baseline. Windows CI is pinned to openpyxl 3.1.5; verify after any dependency update.

## Release gate

Run full Windows Actions Python tests and packaged EXE smoke at the published SHA. Verify EXE, SHA-256 and dependency assets before declaring Release success. Keep all previous tags. Do not start 8-hour soak yet.
