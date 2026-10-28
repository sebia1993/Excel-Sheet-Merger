# Pre-soak stability: fixed-local source path safety

Date: 2026-10-28. Approved scope: continue non-visible stability improvements before scheduling the final 8-hour Codex/Windows test.

## Why
The current `_is_fixed_local_drive()` calls GetDriveTypeW on the drive root and tests `path.is_symlink()` only on the final file. For a Windows junction/symlink **in an ancestor directory**, a file may appear under `C:\` but live on a network share or cloud/reparse-backed path. It is then processed by a full (8–32MiB) hash rather than the deliberately I/O-bounded 17-window scan. The risk is network stalls and unnecessary SMB reads; input data and visible results are not directly changed.

## Alternatives
1. **Selected: reject any reparse point in path ancestry** before allowing fixed-local full-file hashing; use `lstat()` of each component, checking symlink mode and Windows `FILE_ATTRIBUTE_REPARSE_POINT`. Stop on first redirect, fail closed if metadata can't be inspected. Keep existing drive root detection and use 17-window sampling otherwise. Reading reparse metadata does not follow its contents.
2. `Path.resolve()` and then `GetDriveTypeW`: may touch a disconnected share and trigger expensive remote resolution; avoid.
3. All files use sampled hashes: removes local protection improvement; avoid.

## Constraints
- No UI, input formats, sheet names, output naming, logging, DRM/permissions, or workbook serialization change.
- Do not follow any redirection just to classify the drive. Do not depend on admin privileges or a network server in tests.
- Full hash remains only for confirmed fixed local Windows files of at most 32MiB; other large sources remain sampled (17 x 64KiB).
- Small sources <=8MiB retain existing full fingerprint path.

## Acceptance
- Existing file through an ancestor symlink is classified as unsafe for fixed-local full hashing (testable on Linux with a temporary symlink).
- A simulated Windows reparse-attribute directory is detected, even when it's not a POSIX-style symlink.
- A normal path with no symlinks/reparse points stays eligible after the existing drive-root check.
- Metadata lookup errors cause a conservative sampled fallback, not an exception that breaks an otherwise readable source.
- Existing 122 regression tests still pass. Windows CI also checks a true junction when available.

## Final validation boundary
Corporate DRM/SMB, advanced Excel objects, unbounded UI text logs, and 8-hour Windows soak remain separate tasks. Do not claim production soak or universal source-change detection.

## Additional reproduced GUI defect (same pre-soak scope)
The `결과 폴더 열기` callback calls `mkdir()` and `os.startfile()` without handling `OSError`. Permission loss or disconnect can escape into Tkinter's callback error handler instead of a Korean actionable message. Scope-approved repair: catch only these OS errors around folder creation/Explorer opening; keep the button, folder creation, result path, successful behavior, and the existing UI exactly the same. Present a localized error using the existing `diagnose_failure()` mapping, append the technical failure to the existing on-screen log, and keep the GUI responsive. No automatic retry or data modification. Confirm both creation failure and (Windows-only) Explorer failure with red/green regression tests.

## Additional CI release sequencing risk
The Windows workflow publishes a GitHub release with `--latest` after each main-branch build. Without workflow concurrency or checking that the run SHA still matches `main`, two builds from rapid commits could finish out of order and an obsolete build could become the latest release. **Selected**: use a ref-scoped GitHub Actions concurrency group to cancel older in-flight work, and check current main SHA immediately before `gh release create`; stale jobs skip publishing. Alternative (only concurrency) cannot protect a manually rerun obsolete commit. Do not alter the EXE, installed dependencies, or smoke tests. This is a deployment-stability change, not a user feature.
