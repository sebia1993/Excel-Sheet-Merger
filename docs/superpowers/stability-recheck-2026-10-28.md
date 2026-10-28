# Excel Sheet Merger — Second-pass stability re-audit

Source baseline: `v1.3.12` (`ec060c35750fbf93346b087b22cd1d2c172a544b`).
Inspection scope: zero end-user feature/layout changes; previously reported S-01 through S-10.

## Remediated and verified
- S-01 (P0): stale scan cache cleared on failure, source/output folder binding verified; scan→change folder→fail→no output regression.
- S-02 (P1): mixed CSV/Excel CSV data goes to write-only openpyxl worksheets and is streamed into final XLSX ZIP sheet XML; values, strings, formulas, merges, formats and sheet names tested. 60k rows × 8 columns CSV plus small XLSX: **111.3MB peak RSS**, 11.61 seconds on Linux/Python 3.13 (prior audit 252.1MB, 3.75s; memory improves but zip recompression increases processing time). Relative RSS reduction ~56%; performance depends on hardware.
- S-03 (P1): sparse source uses only instantiated cells (`Worksheet._cells`, with openpyxl 3.1.5 pinned); 700th-row style benchmark: **101.3 MB peak RSS / 0.020s** vs prior 215.9 MB / 0.878s (separate synthetic benchmarks).
- S-04 (P1): per-job owner-locked scratch directory and 24h delayed cleanup for *recognized*, unlocked residue. Recovery tests cover active locks, unknown contents, legacy file preservation and normal cleanup.
- S-05 (P2): atomic text-log publication and distinct warning when TXT save fails after XLSX success.
- S-06 (P2): file size, mtime_ns, filesystem identity snapshots; verified before/after read and immediately before publish; tested pre-merge and mid-read changes. CSV full-file validation remains a deliberate extra read to preserve late-appearing CP949 cases.
- S-07 (P2): Tk events processed in 100-event batches, preserving FIFO semantics and completion ordering.
- S-08 (P2): packaged Windows EXE must perform actual headless CSV+XLSX merge and verify result; normal invocation still opens the unchanged GUI.
- S-09 (P2): Windows 2025 and Python 3.12.10 runner, exact dependency versions taken from last successful Windows action; SHA-256 and build dependency manifest shipped.
- S-10 (P3): native Windows smoke and repeated launch/termination test improve coverage. PyInstaller `_MEI` directories left after forced termination cannot be safely deleted in unprivileged app without knowing process ownership; automatic cleanup intentionally *not* implemented.

## Verification evidence
- Baseline: 67/67 original tests passing (Linux/Xvfb).
- Hardened: 80/80 tests passing (Linux/Xvfb) plus bytecode compilation and `git diff --check` (pre-publication).
- Sample mixed 60k×8 CSV + XLSX successfully generated, read, and checked final rows.
- Previous-release source and spreadsheet engine regression tests preserve grouping, spreadsheet type and style semantics, file protection and error guidance.

## Fresh review — remaining risk recommendations (not changed without further approval)
1. **P1 Windows/SMB/DLP real-machine soak test:** the GitHub hosted runner cannot reproduce corporate DRM, network file locking, backup software or AV. On an approved Windows 11 PC test multiple hours with representative CSV/XLSX/XLSM/XLS and SMB shares; record working set, handles, copied output and exit status.
2. **P2 old scratch ownership:** legacy `_병합중_*.xlsx` files from versions prior to the owner-lock protocol remain; indiscriminate cleanup could delete a live process's intermediate output. Inspect their age/ownership manually before removal.
3. **P2 mutation without stat change:** a program replacing a file with exactly the same size, filesystem ID and nanosecond mtime may evade metadata checks; cryptographic hash on all sources would reduce risk at extra heavy I/O cost, especially on SMB. Reassess if adversarial file replacement is a credible threat.
4. **P2 XLSX ZIP structure validation:** openpyxl integrity checks and EXE smoke are effective but native Microsoft Excel opening and recalculation may differ; representative manual Excel opening remains required.
5. **P2 stream writer output vs in-memory writer:** mixed XLSX now performs an extra ZIP copy/recompression; it reduces memory pressure at the cost of more elapsed time and temporary disk storage (at least two staged XLSX packages). On near-full disks, existing safe failure behavior protects original files, but the operation may not complete.
6. **P3 runtime one-file temp extraction:** forced termination of PyInstaller can leave `_MEI*` folders. Do not auto-delete unknown folders; review Windows TEMP accumulation and endpoint security logs during soak testing.
7. **P3 logs from huge group counts:** the Tkinter history text widget is unbounded, though events are now processed in bounded batches and persistent TXT reports preserve details. A future optional cap would change what is visible in the UI and therefore falls outside the strict no-feature-change constraint.

## Release gate
Only claim released status after confirming the GitHub Windows CI run for the **exact hardened commit** passes every step, including packaged EXE smoke, and GitHub Release contains EXE, SHA-256, and dependency manifest. Preserve `v1.3.12` as a fallback.
