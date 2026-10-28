# Excel Sheet Merger — Stability Re-audit (Round 9, Before Soak)

Date: 2026-10-28. Initial baseline main v1.3.24 at commit 61ae6dcbad17f4fb58477d247f1185c9b1e72ecd; merge.py baseline source blob verified equal to local source.

## Reproduced defect
A ZIP-compressed XLSX of <512 KiB containing syntactically valid worksheet XML of >64 MiB was opened successfully by the existing in-memory parser. The application had no expansion guard. A test was RED (expected SourceFileError, but merge succeeded).

## Changes
- Lightweight ZIP central-directory preflight in xlsx_safety.py: 30,000 member ceiling, 1 GiB aggregate expanded ceiling, 256 MiB XML/relationship member ceiling, >400 expansion ratio for any component >64 MiB, and duplicate package member rejection.
- Existing Excel parser is invoked only when preflight passes. CSV/XLS and output semantics are untouched. The input is never modified.
- A filename-aware EXCEL_RESOURCE_LIMIT diagnostic explains safe checking/re-export; it does not assert malicious intent or encryption.

## Verified locally
- Before/after red-green validation of oversized XML, duplicate [Content_Types].xml, and Korean guidance.
- Nine archive safety tests including normal small compressed workbook, preexisting XLSX protection, cumulative/individual XML quotas and ZIP member count.
- Complete local unittest suite: 137 tests, 0 failures, 4 Windows-only skipped; syntax compilation and git diff whitespace check passed.

## Limits and further recommendations
- **P1** 8-hour company Windows 11 + SMB/DRM/AV and real Excel interoperability remains last, per user direction. Hosted Windows smoke does not replace it.
- **P2** ZIP central-directory declarations can be misleading, and even plausible XLSX/XML with moderate compression can require high peak memory. This guard is not a full memory sandbox. Review genuine large customer samples before changing policy.
- **P2** Existing 32 MiB+ / remote source fingerprints sample portions; an unsampled edit may remain undetected. Full SMB hashing has significant I/O cost.
- **P2** Advanced workbook objects (pivots, charts, external links, cached formulas) retain existing limited fidelity.
- **P3** Long-running GUI Text log memory grows because all history remains visible; optimization without truncating content needs further UX engineering.
- **P3** One-file PyInstaller _MEI* remnants after forced termination cannot be deleted without ownership evidence.

## Next step
Do not claim Windows CI or release success until the new commit's workflow and EXE assets are independently checked. Keep previous good release available for rollback.
