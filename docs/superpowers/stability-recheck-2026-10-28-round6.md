# Excel Sheet Merger — Stability Re-audit Round 6

2026-10-28, baseline v1.3.19 (`73d7d503a2543c32ae2b9535b635566006837ec8`). User pre-authorized stability improvements without extra routine approval.

## Reproduced fault and resolution

- **Fault:** a 17MiB source CSV altered at offset 512KiB with identical size, mtime and inode is outside existing 17 sample windows; scan+verify previously missed this modification.
- **Repair:** classify Windows fixed local drives with `GetDriveTypeW`. For 8–32MiB files on confirmed fixed local drives, stream the entire source into the original BLAKE2b fingerprint. For all other large paths (SMB/UNC/mapped/unknown or >32MiB), keep the 17 x 64KiB scheme. Existing <=8MiB full hashing retained.
- No user-facing features, workbook contents, file naming, DRM/permission handling or UI code changed. No added dependency.

## Test-first evidence

- Failing case observed before change: `test_confirmed_local_drive_detects_stealth_edit_between_old_samples` failed with `AssertionError: SourceFileError not raised`.
- After change: Linux local Xvfb full suite **115 tests, OK (2 Windows-only skipped)**. Windows-only fixed-drive and mocked remote-drive tests require Actions validation.
- Local buffered I/O synthetic benchmark (median, 4 runs): 17MiB sampled 0.0016s vs complete 0.0250s; 32MiB sampled 0.0017s vs complete 0.0477s; 33MiB both sampled 0.0015s. Local Linux host, not SMB or Windows results.

## Remaining issues and guardrails

1. **P1** Real corporate SMB/DRM/AV + 8-hour soak cannot be verified on GitHub hosted runner. Use approved sanitized Windows fixtures.
2. **P2** Unsampled changes in remote, unknown, or >32MiB files still evade detection; full hashing everything risks large repeated network I/O. Investigate OS-assisted change notifications or optional local-only verification if real incidents arise.
3. **P2** Complex workbooks with images, pivots, macros, formula calculation, and external connections are not faithfully copied; test with actual Excel for business usage.
4. **P2** Save response uncertainty on unstable SMB: do not automatically retry, inspect result folder.
5. **P3** Long session UI log text memory remains unbounded because user wants all visible records; avoid truncation without design approval.

## Release gate

Only mark Windows verification and Release successful after Actions is green on the exact new main commit and the new `ExcelSheetMerger.exe` plus SHA256 checksum are attached. Preserve previous stable Release.
