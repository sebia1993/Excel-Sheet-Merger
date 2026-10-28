# Excel Sheet Merger — Windows CI and Network-Share Stability (Round 5)

Date: 2026-10-28. Authorization: user explicitly requested successive stability hardening without feature/UI changes and gave standing consent for ordinary GitHub commits and releases.

## Purpose
Prevent silent partial scans and unclear failures caused by transient SMB/network outages, while modernizing GitHub Actions JavaScript runtimes. Preserve all existing file formats, worksheet contents, UI layout, success-path behavior, naming rules, and policy-safe diagnostics.

## Alternatives and decision
- **Chosen: fail closed for explicit network errors, no implicit retry.** Network storage can have ambiguous commit outcomes (e.g. server created output just before connection dropped). Identify specific OS/Windows codes, reject a scan when access to source data failed due to disconnection, keep existing files untouched, and advise checking the output folder before retrying.
- Blind automatic retry: rejected, because final output may already exist; it also masks ongoing problems and duplicates data.
- Treat every I/O error as SMB: rejected, because plain local disk and permissions failures have different remedies.
- Leave existing GitHub Actions majors in place: rejected, because GitHub removed Node 20 from hosted runners on 2026-09-23.

## Implementation boundaries
- Use a small, pure helper `is_network_io_error()` for recognized `errno` / Windows `winerror` conditions, including nested failures; do not inspect localized error messages.
- `classify.scan_folder()` retains DLP/PermissionError handling for fingerprints, but explicit network failures must terminate the scan rather than accepting a half-verified group.
- `failure_guidance.diagnose_failure()` returns safe Korean guidance when the error chain contains a network failure. The existing scan error dialog shows that guidance if scanning fails due to a disconnection. The existing UI error banner and text log remain unchanged structurally; never assert that corporate DRM/SMB is certainly at fault.
- `merge_group()` remains fail-closed without retries. A transient read/write error must not publish a partial workbook; uncertainty after a final network publish must be disclosed rather than assumed unsuccessful.
- CI Windows runner remains `windows-2025`, Python 3.12.10 x64, pinned Python build dependencies, PyInstaller command, and EXE smoke test. Replace first-party Node 20 action generations with checkout v5, setup-python v6, upload-artifact v6; requires hosted runner >= v2.327.1.

## Acceptance and checks
1. Explicit SMB errors during fingerprint/scan do not leave a mergeable partial report; ordinary permission/DLP failure remains handled as before.
2. Network failures from source reading, final staged write, or final publication produce a Korean network/sharing diagnosis and retain original workbooks. Do not claim the destination file was definitely absent if publication outcome is ambiguous.
3. Non-network `EIO`, `PermissionError` and `ENOSPC` do not get misdiagnosed as SMB errors; preserve existing diagnostic precedence.
4. Existing normal CSV, XLSX, XLSM, XLS merge output and UI tests continue to pass.
5. All three GitHub Actions support Node 24; the full Windows suite, bundled EXE smoke and release assets verify against the exact tested commit.
6. Corporate SMB/DRM interruption and 8-hour soak remain real-PC acceptance gaps, not claims of this round.
