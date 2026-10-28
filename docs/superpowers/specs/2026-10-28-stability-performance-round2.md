# Excel Sheet Merger stability / performance follow-up

Date: 2026-10-28. User approval: continue existing stability improvements without changing features, file naming, UI or formats.

## Outcome

Reduce the redundant CPU cost of ZIP recompression in mixed CSV+Excel merges without compromising file compatibility, contained CSV text, source styles/formulas, existing filenames, or resource safety. Add bounded CI workload verification for repeated runs, rather than claiming production 8-hour / corporate-DRM coverage.

## Alternatives considered

- **Selected**: keep the existing streamed CSV + ZIP XML stitch and tune the streamed worksheet member's DEFLATE effort only. Preserve other members' metadata and compression. Benchmark the time / final size trade-off, ensure CRC and XLSX parsing.
- Full rewrite to stream complete mixed workbook: overlarge, risk style and formula changes.
- Transfer compressed ZIP members with private internals: faster in theory but brittle and harder to validate on Windows.
- Compress streamed XML at level 1: faster but the benchmark shows much larger results; poor space trade-off.
- Leave the algorithm unchanged: lowest code risk but known excessive CPU overhead for repeat office jobs.

## Constraints

- Visible behavior unchanged: same CSV/XLSX/XLSM/XLS handling, time-named worksheets, checkbox selection, output suffixing, errors, logs, Excel cell types/styles, Korean office UI.
- Legacy multi-workbook, macro handling, source validation, and scratch protections remain untouched.
- Do not open or modify user/company files; fixtures are synthetic; no DRM workaround.
- Follow test-first: measure baseline against a deterministic fixture, add ZIP/XLSX output integrity regression and repeated-workload tests, apply small scoped change, re-measure.
- Only publish new release after Windows CI tests and actual packaged-EXE smoke pass; keep old release available.

## Acceptance

1. Same output filenames, worksheet order and data, including text formulas vs real Excel formulas, merged cells and styles.
2. Export ZIP CRC check succeeds; streamed sheet content remains readable and correct.
3. In repeatable local benchmark, stitched ZIP step materially faster than original level 6 with no disproportionate output growth (target <=10% increase with representative data).
4. Windows CI full suite and EXE smoke pass, no regression to existing no-overwrite and safe cleanup.
5. Report remaining corporate Windows/SMB/DRM validation gaps and newly discovered risks honestly.
