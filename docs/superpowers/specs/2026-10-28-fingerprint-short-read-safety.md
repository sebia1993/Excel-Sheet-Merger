# Excel Sheet Merger — incomplete source fingerprint read safety

Date: 2026-10-28. Scope: authorized stability hardening only, no UI or merge feature changes.

## Observed gap
`file_fingerprint(path, size)` trusts a read even when the bytes actually read are fewer than the size from `file_snapshot`. This can produce a fingerprint for an incomplete source stream. In bounded sampling, an individual seek/read can also return fewer than the expected 64 KiB without being rejected. On an unstable filesystem, this may accept an incomplete observation as valid.

## Approaches
- **Selected:** require the complete byte count in full-hash mode; for sampled windows, collect short reads until exactly 64 KiB or raise a recognizable `ValueError` on unexpected EOF. Preserve the current BLAKE2b digest algorithm and 17 offsets. A truncated, unstable file cannot be represented by a valid content fingerprint.
- Always full hash large and remote files: rejected as it expands SMB reads dramatically and still cannot prevent mid-read network failure.
- Ignore short reads and rely on after-the-fact metadata only: rejected because remote cached metadata can be stale.

## Constraints
- Preserve all current classifications, normal CSV/XLSX/XLSM/XLS output, names, styles, checkbox UX and file-protection policies.
- Never modify company or source files; use synthetic tests.
- Keep 17 x 64 KiB bound on sampled remote/large sources and streaming 128 KiB chunks for complete hashing.
- Do not treat legitimate unavailable DRM fingerprint access as proof of corruption; existing PermissionError handling must remain unchanged.
- On a detected truncated read, fail closed: scan must not publish a validated group based on incomplete bytes and merge must not publish output.

## Acceptance
- Both full-hash and 17-window modes reject short/incomplete reads with clear error, even when the caller supplied a stale expected size.
- Sample reads that return partial chunks but eventually complete can still compute the original fingerprint correctly.
- Normal CSV/Excel mixed merge and existing 115 tests continue to pass, with new regression cases.
- Windows hosted tests and PyInstaller EXE smoke pass before publishing; no claim of real corporate SMB/DRM soak verification.
