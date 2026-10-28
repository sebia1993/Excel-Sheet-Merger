# Stable local-source fingerprint design (2026-10-28)

## Intent and constraints
The 17-window BLAKE2b fingerprint protects against many metadata-preserving edits but misses modifications in unsampled regions of large files. The user authorizes continued stability hardening, not feature/UI changes. Keep CSV/Excel loading, group selection, name numbering, error text and sharing policies unchanged.

## Options considered
1. **Selected:** full streaming fingerprint for 8-32 MiB files on confirmed Windows fixed drives only; retain 17-window, ~1.06 MiB read budget for SMB/UNC/mapped drives and files >32 MiB. Windows GetDriveTypeW is used conservatively. Unknown, relative/UNC and symlink paths remain sampled. No extra packages. This closes a real local-file blind spot while avoiding full SMB file reads.
2. Always full hash: strong detection but repeated 32 MiB to multi-GiB network traffic per scan and source validation, unacceptable SMB performance risk.
3. More sparse samples: low I/O but no guarantee for a specific unseen byte between samples; still misses observed gap.

## Expected behavior
- For a known fixed Windows drive and 17MiB CSV, a same-size/same-mtime edit between current sample offsets raises a source-change error before a result is published.
- Small files <=8MiB use full hashing as before. Unknown path, SMB/UNC, mapped network drive or >32MiB file keep the 17 windows.
- Keep each stream read bounded at 128KiB; on Windows do not inspect file contents via OS APIs, and do not copy user data anywhere.
- Preserve worker UI, data output, temporary-file cleanup, existing error handling and packaging.

## Verification
- TDD failing case: 17 MiB sparse file mutated at 512KiB with restored st_mtime and unchanged size/inode; selected local drive full mode must detect before any final result.
- Read-budget tests prove remote/large file reads capped; unit tests for symlinks/UNC/drive type fail-closed.
- All regression tests, local benchmark, Windows CI, packaged EXE smoke; claim no corporate SMB performance validation.

## Remaining caveat
For SMB, very large files and unrecognized drives, sample gaps remain. This is a bounded-risk performance policy, not proof of all-byte integrity.
