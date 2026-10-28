# GUI Stability Follow-up — approved scope

Status: user authorized continued stability work on 2026-10-28. Baseline: public release v1.3.13, remote commit `4d77b0134e288a4811f65fbc5b9fc273a5dfa895`; local source `a212d6e` matches the key Git blob hashes.

## Goal

Protect the existing Excel Sheet Merger UI from exceptional inability to start a worker thread; lower Tcl/Tk calls for log-heavy jobs without changing visible results or existing merge functionality.

## Non-goals

- No CSV/XLS/XLSM/XLSX parsing change, workbook content or filename change, screenshot/layout change, checkbox or progress changes.
- No automatic deletion of unknown scratch files, DRM bypass, corporate-system access, or unsupported background processing.
- Do not discard historical GUI log lines or truncate the persistent TXT record.

## Candidate approaches

- Recommended: recover worker startup errors synchronously in `_start_merge`, then batch consecutive log queue events into one text-widget insert per dispatch; retain timestamps and order.
- Alternative: limit text-widget history or change log view to a virtualized table; rejected for this iteration as altering observable history and requiring larger changes.
- Alternative: auto-retry thread creation; rejected as potentially amplifying resource pressure and hiding a real failure.

## Behavior contract

1. Starting a merge with valid groups normally still starts exactly one background worker and disables input while busy.
2. If creation or starting of worker raises RuntimeError or OSError, no background job is running; GUI is usable again, error is surfaced in Korean, and no workbook is created by that attempted call.
3. The next ordinary merge can still be started. Group checkbox state, source scan, result folder, progress format remain unchanged for successful runs.
4. All queued log messages remain visible in the same order with the same timestamp format and persist under the unchanged job record rules. Failure and done events retain FIFO ordering.
5. A burst of 100 log events should result in one Tk text insertion rather than 100. UI event queue continues to yield every 100 processed events as before.
6. Original scan, merge, error advice and scratch protection contracts remain intact; all 80 existing baseline tests must pass.

## Acceptance and review

Write GUI behavior tests first and watch them fail; implement minimal changes in `gui.py`; run targeted and full Xvfb unittest suite, compileall, and smoke checks. GitHub Actions Windows test, packaged EXE smoke and Release on the exact commit must succeed before claiming publication. Re-audit afterward and report any remaining windows-specific or corporate DLP/SMB limitations.
