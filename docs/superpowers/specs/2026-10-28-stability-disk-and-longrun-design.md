# Excel Sheet Merger — low disk capacity and long-run stability

**Date:** 2026-10-28. **Approval:** ongoing "진행해" instruction for reliability-only improvements. Existing UI, input formats, output names, grouping, workbook cell values and style semantics must not change.

## Problem and evidence

The mixed CSV+Excel pipeline writes `base.xlsx`, `csv.xlsx` and `result.xlsx` under the result drive in a private, locked scratch directory. If the result drive runs out of space during the final ZIP construction, the OS raises an I/O error and the previous result remains safe, but that error is currently classified as an unknown generic failure and low capacity isn't detected before expensive stitching. Accurate peak space is hard to predict because CSV->worksheet XML changes size and compression ratios. A preflight **estimate** is a safeguard, not a guarantee or a reservation.

The existing GUI Text control retains all session lines by design. Automatic truncation would change the accessible UI log, so this round measures long-session memory and event processing but **does not delete or hide old log entries**. TXT records are fully preserved.

## Selected approach / alternatives

- **Selected:** when the intermediate archives already exist, estimate minimal additional free space based on their ZIP sizes and a modest safety margin. Query free bytes on the *actual result drive*. If clearly insufficient, stop before building `result.xlsx` with an actionable Korean disk-space error. If disk capacity lookup is unavailable (notably some SMB shares), let normal safe I/O continue; never guess. Classify both OS `ENOSPC` and quota `EDQUOT` in existing failure guidance. Make the guard directly testable, and leave the atomic rename/no-clobber logic untouched.
- **Rejected:** move scratch to the system TEMP drive; atomic publishing would cross filesystems, undermining no-overwrite guarantees.
- **Rejected:** stream ZIP internals with private APIs or allocate all CSV data in RAM; high functional/compatibility and memory risk.
- **Deferred:** silently prune the Tkinter log or implement virtualized history; changes existing user-visible browsing of session logs.

## File boundaries

- `src/excel_sheet_merger/merge.py`: lightweight non-destructive capacity estimate between writing two staging archives and the final ZIP stitching; best effort skip if disk free-space metadata inaccessible; preserve exceptions and cleanup.
- `src/excel_sheet_merger/failure_guidance.py`: specialized Korean OS disk-full / disk-quota explanation, scoped to error codes in the exception chain, without conflating DRM and encoding errors.
- `tests/test_disk_capacity.py`: low free, enough free, unknown free, simulated ENOSPC on write and existing result protection; real mixed workbook assertions.
- `tests/test_failure_guidance.py`: actionable disk/EDQUOT classification.
- `benchmarks/measure_round3_resources.py`: optional synthetic event/memory measurement (no corporate data, no GUI history truncation). Logs and summary document the limits; no eight-hour corporate-PC claim.

## Constraints and acceptance

1. CSV-only path remains unchanged; mixed path creates the same XLSX workbook and no altered cell types, styles, sheet names, or filenames.
2. A clearly insufficient result filesystem triggers `OSError(ENOSPC)` **before** final ZIP stitching, never publishes a partial XLSX, cleans all private scratch files and preserves existing outputs.
3. If `disk_usage` is unsupported/unreliable, do not block normal processing. Remaining races and quota changes may still cause later `ENOSPC`; handle them with the same error guidance.
4. No new Windows runtime dependencies, same standalone EXE build and smoke-test procedure. Company/DRM/SMB/AV soak requires approved host and cannot be simulated fully in GitHub runners.
5. Retain all in-session GUI text logs, measure long-run overhead, and report findings rather than unilaterally changing history retention.
