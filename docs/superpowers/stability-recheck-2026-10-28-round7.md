# Excel Sheet Merger — Stability Re-audit Round 7

Date: 2026-10-28. Baseline: v1.3.20, `9b833f32d0ddec90c52a90d8c3fe857c8d18f110`.

## Defect found and corrected

`classify.file_fingerprint()` previously returned a successful digest after unexpected EOF: for full hashes, it did not compare total bytes read to the stat-recorded size; in 17-window large-file mode, it accepted a single short `read(65536)` as a complete sample. This could cause incomplete/untrustworthy file data to be accepted after a filesystem interruption or stale metadata read. A stream that returns a transient nonempty short read could also hash a partial window instead of the actual full 64KiB window.

- Changed full-file fingerprinting to verify total bytes read exactly equals expected size and reject unexpected growth or EOF. Still reads in 128KiB chunks.
- Changed 17-window fingerprinting to finish each exact 64KiB window with bounded repeated reads, rejecting premature EOF; 17 positions and 1.0625MiB maximum total bytes read are unchanged for normal inputs.
- Preserved the file name in `SourceFileError` for incomplete fingerprint checks during the final publication validation.
- Added a specific Korean warning for incomplete read with plausible source-change or share interruption causes, without claiming DRM encryption.
- Left actual workbook transformation, folder/UI interaction, output naming, DRM policies and read-only source behavior unchanged.

## Test-first evidence

- Before implementation, 4 new tests failed, including false success on a 100-byte reported source with only 8 bytes available, false success on a 9MiB reported source with only one window available, a split 64KiB window producing a different digest, and a scan accepting a stale reported file size.
- Additional tests reproduced missing `SourceFileError` file name and unhelpful `UNKNOWN` diagnostic during incomplete read.
- After implementation: local Linux/Xvfb `python -m unittest discover -s tests -v` **122 tests OK (2 Windows-only skipped)**; `python -m compileall -q src tests` passed. Real Windows and packaged executable checks are required after GitHub publish.
- Synthetic tests do not demonstrate actual corporate SMB, DLP, DRM, antivirus or 8-hour operational behavior.

## Residual risk priority

1. **P1:** Enterprise SMB/DRM/AV/VDI environment and eight-hour Windows soak still need authorized real-device testing.
2. **P2:** Large files on network or >32MiB source still use 17 sampled windows; edits outside windows can escape detection with unchanged metadata. Full hash costs additional SMB I/O.
3. **P2:** A network server can acknowledge final publication and then lose its connection. Automatic retry could create duplicates; user must confirm the result directory first.
4. **P2:** Complex Excel objects (charts, pivot tables, cached formula evaluation, external links) are not guaranteed to be copied completely; validate real authorized examples in Excel.
5. **P3:** Long-lived full Tk log history consumes memory; preserving all log lines precludes simplistic truncation.

No company files accessed, no user file deletion, no new dependencies. Older Releases must be preserved.
