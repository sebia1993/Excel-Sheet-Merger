# Excel Sheet Merger — Stability Follow-up Audit

Scope: release v1.3.13 base, preserve all end-user behavior and UI design. New fixes validated on an isolated worktree.

## Fixes and evidence

1. **Worker thread startup failure (newly reproduced):** baseline `_start_merge` sets `busy=True` before `thread.start()`, but a RuntimeError/OSError during construction or start leaves the interface stuck disabled with no worker to produce a completion event. Regression tests reproduce both cases; fix restores action states and provides error guidance while keeping the previous scan and checkmarks. New normal merge can start afterward.
2. **Log burst GUI responsiveness (newly measured):** 100 consecutive `log` events previously resulted in 100 individual `Text.insert` calls; buffered consecutive events now require one call for the same log content. FIFO boundaries for progress, failures, completion are preserved. Displayed history and stored TXT are not truncated.
3. **CSV+XLSX repeated stress:** 60 sequential mixed merges with fresh numbered results; every output workbook opens and retains `11시` and `14시`, source CSV `00123`, and Excel formula `=1+2`. Local test approximately 3.3 seconds, RSS ~101.4 MiB baseline → 104.6 MiB peak; scratch folders left: zero. A synthetic Linux/CPython 3.13 check, not an 8-hour Windows test.
4. **Regression:** baseline 80 tests; follow-up 84/84 pass under Linux Xvfb; bytecode compile succeeds, git diff whitespace check clean. Windows CI EXE test/release status must be verified against the exact GitHub commit before claiming release completion.

## Second re-audit: remaining candidates (no automatic feature changes)

- **P1: corporate Windows/SMB/DRM/endpoint security**: current CI uses an isolated hosted Windows VM, cannot reproduce corporate document security policies and network share latency. Run a user-approved on-site 8-hour soak with representative sanitized data; capture working set, handle count, stop/continue behavior and file integrity.
- **P2: large mixed workbooks consume additional scratch disk and CPU**: ZIP recompression after streamed CSV lowers peak memory, but increases I/O. Benchmark on the target hardware and free disk margin before changing the algorithm.
- **P2: metadata-only source snapshot check**: same-size/same-mtime in-place modification could evade detection; full cryptographic hash requires extra I/O. Consider a configurable verification policy only after performance measurements and agreement on behavior.
- **P2: Excel advanced object fidelity**: chart, pivot, images and calc engine are not guaranteed. Representative files must be opened by a native Excel version for acceptance; avoid promising full fidelity.
- **P3: visible log widget unbounded by design**: batching improves responsiveness but cannot cap RAM if history accumulates indefinitely without changing visible history. Investigate virtual display only if a genuine hours-long workload reveals growth.
- **P3: old scratch / `_MEI`**: do not indiscriminately delete unowned legacy or PyInstaller temporary files; require owner checks and security-policy compliance.

## Safety and rollback

Keep the previous v1.3.13 Release downloadable. Do not automatically start an 8-hour task or claim one ran. All changes are scoped to GUI defensive error handling, batched log rendering, tests and documentation; the merger, classifier, outputs and theme remain byte-identical to the baseline.
