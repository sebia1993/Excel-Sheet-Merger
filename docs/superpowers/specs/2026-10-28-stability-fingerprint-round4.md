# Excel Sheet Merger — Large-file fingerprint stability round 4

Date: 2026-10-28. User authorized ongoing stability improvements without changing any user-facing functionality and requested approval only for genuinely blocked operations.

## Goal
Reduce the blind spots of metadata-preserving edits in inputs larger than 8 MiB, while keeping network-drive IO bounded and preserving exact GUI, import, output, naming, and security behavior.

## Existing mechanism
The input scan records size, nanosecond mtime, device/inode and a BLAKE2b fingerprint. For <=8 MiB, it hashes the complete file. For larger input, it reads three 64 KiB regions (head, midpoint, tail), then the merge verifies those same regions before and after processing and before publishing. An edit outside the sampled ranges that restores metadata may evade verification. No security guarantee is claimed from sampling.

## Approaches
- **Selected:** evenly spaced 17 sample regions for files >8 MiB; keep existing full-file behavior for small files, bounded 1.0625 MiB per check for large files (17*64 KiB), with offsets incorporated into digest. This improves coverage across the entire source with bounded reads and modest peak memory; still cannot detect changes in unobserved regions.
- Full-file content hash for all files: stronger detection, but repeatedly reads every large source on scan, before, after, and on publish. Unacceptable potential SMB latency and load without real corporate network benchmarks.
- Keep three regions: zero extra IO, but known blind spots make protection weak against arbitrary mid-file edits.

## Constraints
- Do not change input types, group naming, time-based sheet names, CSV text, Excel cell/style/fomula, checkbox selection, UI, or result suffixes.
- Preserve enterprise DLP behavior: read failures at scan are not silently classified as rejected by fingerprinting; merge-time errors still display source filename and advice.
- Memory remains bounded (one 64 KiB chunk at a time); no new dependencies and no full-file buffering.
- Do not treat fingerprints as a cryptographic guarantee for unsampled regions. Clearly report remaining limitation.

## Acceptance
- A large file whose early-quarter bytes change with identical size, mtime and inode is rejected before publishing.
- Also detect a mutation at a later fraction outside original head/center/tail windows.
- Unchanged large sources pass; below-threshold files retain complete read semantics.
- The large-file fingerprint makes only 17 bounded reads, <=17*64KiB total; no full file scan.
- All existing 99+ regression tests pass and output ZIP, data, and naming remain stable.
- GitHub Actions Windows and packaged EXE must pass before claiming publication is successful.
