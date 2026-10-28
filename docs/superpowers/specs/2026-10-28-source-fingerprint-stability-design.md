# Source Fingerprint Stability — approved continuation

**Date:** 2026-10-28. **Baseline:** GitHub v1.3.14, commit `48bf12c404e663b333c63c5dd17c754ef7554186` (same tested files as local `cce38db`).

## Goal

Close a known source-consistency gap: an input may change after analysis with the same byte count and restored modification time. Prevent publishing results derived from that input when its inspected bytes change, without changing the CSV/Excel business behavior, GUI, or output filename rules. Protect the existing user experience for blocked corporate files.

## Alternatives considered

- Full SHA-256 on every scan and every pre/post merge check: exhaustive but costly for large SMB/enterprise files and freezes the synchronous scan on gigabytes. Rejected as current default.
- **Selected: bounded content fingerprint + metadata**. Hash the full contents of input files up to 8 MiB and three 64 KiB spans (start/middle/end) for larger files. Verify on scan, before read, after read, and immediately before publishing. Low memory and predictable I/O, detects changes in sampled ranges and all changes in small files; misses unsampled middle changes in large files.
- No further checks: fastest but leaves same-size/same-mtime input mutation undetected. Rejected.

## Behavioral contract

- `file_snapshot` retains its exact 4-tuple API; add `content_fingerprint: bytes | None` to `SheetInput` only as internal scan state, and `file_fingerprint(path, size)` in `classify.py`.
- When scan can read the input: compute fingerprint; ensure pre/post metadata match, fail scan for inputs changing during analysis.
- When OS denies fingerprint read on a scan that otherwise succeeds, leave fingerprint unset and preserve the previous behavior of reporting read/DRM failure during merge; never circumvent DLP.
- A selected input whose fingerprint no longer matches fails with `SourceFileError` before creating final XLSX. Other groups continue. Preserve original input data, existing XLSX, group checkboxes, names and errors.
- Large-file sampling costs at most 192 KiB per check plus seeks. Full-file checks for <=8 MiB. Remote SMB seek latency may still be significant; benchmark and document.
- Never claim that sampling detects arbitrary same-size/mmtime modifications of files >8 MiB; treat this as a residual P2 risk.

## Acceptance tests

1. CSV modified after analysis with identical size and `mtime_ns` is rejected (previous implementation allows it).
2. Large (>8 MiB) file mutated near middle (a sampled region) with unchanged metadata is rejected.
3. Unchanged CSV/Excel inputs still merge, preserving `11시`/`14시`, values and formulas.
4. Fingerprint I/O problems under scan do not suppress existing merge-time DRM/error guidance.
5. Full baseline 84 tests remain green, and new tests prove the regression.
6. GitHub Actions on Windows tests and packages exact committed source. Built EXE real merge self-test remains green. No deployment claim before verifying assets.

## Explicit limitations

Does not implement a full cryptographic audit trail, guaranteed consistency against active same-mtime concurrent edits during a single read, access through corporate security software, deep Excel shape fidelity, or 8-hour Windows corporate soak.
