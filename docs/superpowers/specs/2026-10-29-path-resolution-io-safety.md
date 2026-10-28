# Folder Resolution Errors during Scan and Start — Stability Design

Date: 2026-10-29. User delegated stability-only fixes and distribution; no per-step approval needed. Company Windows 11 eight-hour soak remains the final stage.

## Reproduced path and risk

`MergeApp._scan()` scans a folder inside a try block, then calls `Path(location).resolve()` after that try block. A disconnection/permission error during resolution raises from a Tk callback and may leave a newly assigned `self.report` with stale `self._scanned_folder`. The merge-start path similarly calls `Path(self.folder.get()).resolve()` outside an exception handler. A lost share or redirected path resolution error leaves an old preview actionable and propagates a GUI callback error. No output must be generated with mismatched snapshot/path state.

## Chosen behavior

- Resolve folder path in `_scan()` inside the existing catch and **before** committing a report to the UI. If scanning or resolution fails, the old report and controls remain invalidated.
- If path resolution fails during `_start_merge()`, invalidate stale report/checked groups and show an actionable error rather than starting a worker. Give the affected path; network failures may use existing network guidance. Keep the current mismatch warning for successfully resolved but changed folders.
- Catch only path resolution OSError / RuntimeError. Do not broadly swallow arbitrary programming bugs. Preserve every success path, group selection, CSV/XLSX/XLSM/XLS handling, existing logs and output filenames.

## Tests

A. Completed scan followed by `Path.resolve()` OSError at merge start: no callback crash, no worker started, old preview cleared, buttons disabled, user-friendly warning.
B. A valid newly scanned folder but a `Path.resolve()` OSError at scan completion: no half-committed report, no output or worker, error popup, rescan can recover.
C. RuntimeError resolution (e.g. link loop) is safely rejected during merge start; ordinary matching folder still starts.

## Non-goals

No retry loop, no background scan redesign, no new runtime dependency or UI layout changes. No 8-hour stress test yet.
