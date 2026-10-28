# Excel Sheet Merger — 추가 안정성 감사 (source fingerprint)

**Scope:** v1.3.14 release, same local source Git blob hashes as remote commit `48bf12c404e663b333c63c5dd17c754ef7554186`. Source behavior unchanged for valid inputs. No changes to UI layout, output file names, or conversion engine.

## Fixed/reproduced

1. **Source data changed with identical file metadata:** previous `file_snapshot` checked only size, mtime_ns, dev, and inode. In a synthetic CSV edited in place with the same byte count and restored original mtime_ns, the old code published a stale-scan result. New small-file content digest detects it and prevents publication. Test observed FAIL on baseline and PASS after fix.
2. **Large sampled file changed near middle:** previous code failed to notice same-metadata mutation in an 9 MiB file. New digest of start/center/end 64KiB chunks detects it; FAIL→PASS observed. Above 8MiB this is only sampled verification, **not a full data integrity proof**.
3. **Clear Korean remediation:** existing `failure_guidance` treated detected input changes as `UNKNOWN`, even though the root cause was known. The new `SOURCE_CHANGED` guidance explicitly instructs re-analysis. FAIL→PASS observed.
4. **Scan blocked by DLP/permissions:** fingerprint read failures of type `OSError` do not remove the otherwise-named input from groups; merge-time access/DLP error handling remains unchanged. No protection bypass.
5. **Windows distribution:** packaged EXE self-test count increased from 3 to 10, retaining other CI gates. Windows Actions must pass and release assets must be checked before publication claims.

## Measured performance and resource impact

Synthetic Linux/Python 3.13, local filesystem, openpyxl 3.1.5, mixed workbook with 20,000 CSV rows (351,320 bytes) and one XLSX containing a formula; outputs verified for sheet names/values/formula. Separate baseline/fixed processes:

| Case | Baseline v1.3.14 | Fingerprint version |
|---|---:|---:|
| 30 folder scans | 0.0039 s | 0.0226 s |
| 15 consecutive mixed merges | 7.0686 s | 7.8039 s |
| Peak process RSS | 105.83 MiB | 105.83 MiB |

**Interpretation:** 0.7s higher for 15 merge operations (~10.4%); protects accidental source changes but costs I/O. A synthetic 128MiB sparse input: 50 sample-only fingerprint checks took ~0.017s and each check reads at most 192KiB. These numbers do **not** model SMB latency, DRM, Windows Defender scanning, or typical company hardware. Do not present as guaranteed speed.

## Regression and packaging evidence

- Existing baseline: 84/84 local tests under Linux Xvfb.
- Updated: 89/89 local tests under Linux Xvfb, Python compilation and whitespace checks.
- Windows GitHub Actions and Release outcome: verify after publishing; independent Windows EXE tests must still succeed.

## Remaining risks (post-fix second review)

- **P1 (external validation):** corporate Windows SMB/DRM/DLP, antivirus and 8-hour soak are not reproducible on hosted CI and need an authorized real workstation.
- **P2 (integrity):** >8MiB input is sampled; an unsampled same-metadata edit can evade the check. Full SHA-256 on large inputs would require more I/O and may freeze synchronous GUI scanning on SMB; benchmark against real workload before opting in.
- **P2 (performance):** streamed CSV+Excel mixed merge uses temporary ZIP files and ZIP recompression, trading CPU/disk for bounded memory; no untested ZIP optimization in this iteration.
- **P2 (Excel compatibility):** macro, image, chart, pivot, external workbook references are not fully preserved by openpyxl, unchanged behavior.
- **P3:** visible GUI Text history grows indefinitely during long unattended runs; render batching helps speed but not unbounded memory. Determine real multi-hour workload requirements before truncating or virtualizing.
- **P3:** unknown pre-v1.4 scratch and PyInstaller `_MEI*` folders must not be blindly deleted. Respect company file/endpoint policies.

## Reproducibility / rollback

Check source, tests, build job and original EXE Release. All source code changes scoped to `classify.py`, `_verify_source` in `merge.py`, and `failure_guidance.py`; GUI, theme, parser, file naming and business logic intentionally untouched. Prior releases remain available.
