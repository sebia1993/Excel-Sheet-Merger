# Excel Sheet Merger — Stability Re-Audit (Round 2)

Date: 2026-10-28. Baseline: v1.3.14 source / 84 passing tests; integrated onto newer source-fingerprint commit 7302823.

## Changes implemented

1. **Mixed workbook ZIP recompression CPU:** normal Excel members remain byte-content/metadata equivalent. Streamed CSV XML is recompressed with DEFLATE level 5 instead of the default level 6; this avoids the major file-size increase seen at level 1 or 3. No GUI, CSV or Excel business logic changed.
2. **ZIP compatibility regression:** direct test checks ZIP CRC, no duplicated archive members, exact CSV worksheet XML, untouched member metadata, formula/style/merge-cell retention and CSV string types.
3. **Repeated output integrity:** automated 100-cycle CSV+XLSX merge checks name suffixes up to `(100)`, preserves the original file, validates opened workbooks periodically, and checks for scratch cleanup.
4. **Performance benchmark:** synthetic 30,000-row worksheet: original level-6 0.346 s / 1,658,031 bytes; optimized 0.210 s / 1,757,278 bytes (39% faster within stitch, ~6% larger). Separate synthetic 60,000-row full conversion median 5.011 s → 4.931 s, total ~1.6% faster; output size ratio 0.999. Results are Linux/CPython 3.13, not corporate Windows measurements.
5. **Validation:** baseline 84 passing tests, updated local suite 86 passing tests under Linux Xvfb; integrated suite 91 passing tests with the concurrent fingerprint fix. Windows Actions and packaged EXE must be verified against exact published SHA before claiming Windows success.

## Re-audit — items intentionally deferred

- **P1 corporate-environment soak:** test 8 hours on an approved Windows 11 computer and authorized, sanitized SMB/DRM/AV sample files. GitHub Actions hosted runners cannot replicate corporate policy. Do not claim to have completed this test.
- **P2 source identity coverage (partly addressed by concurrent update):** new fingerprint verification catches same-size/same-mtime changes in files up to 8MiB, and samples beginning/middle/end of larger files. Changes outside the sampled ranges could still escape detection; hashing large files in full would increase SMB I/O. Verify on approved shared-folder fixtures.
- **P2 scratch free-space risk:** mixed merges currently need temporary base, streamed CSV and final ZIP at once. Low disk space results in a failure with cleanup, but a preflight estimate is not available. Benchmark worst-case disk peaks with user-scale files before implementing a possibly misleading size check.
- **P2 advanced workbook content:** complex Excel objects (charts, images, pivot tables, links and cached calculation results) are outside fidelity guarantees. Use native Excel and representative, permitted files for acceptance.
- **P3 long-lived GUI log:** full visible logs are preserved by policy, so Tkinter Text memory can grow without a cap during extremely long sessions. Any truncation/virtualization would change current behavior and needs separate approval; no proven incident yet.
- **P3 PyInstaller `_MEI` cleanup:** cannot safely delete unrelated processes’ private extraction directories. Use OS-supported cleanup and ownership checks only.
- **P3 internal ZIP member compression setting:** optimized writer uses CPython `ZipInfo._compresslevel` because public ZipInfo does not offer a per-member compresslevel argument; Python 3.12 pinned on CI and direct ZIP/XLSX tests protect against upgrade regression. Consider a public-API-only redesign if moving to a different Python runtime.

## Scope & rollback

Preserve v1.3.14 and any interim releases. Do not change functionality, UX, file naming, or security policies. Audit findings are evidence-based; no claim that a short CI test replaces long-running corporate-PC validation.
