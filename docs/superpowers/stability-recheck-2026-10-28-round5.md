# Excel Sheet Merger — Round 5 Stability Re-Audit

Date: 2026-10-28, source baseline: v1.3.18 / 099c4593410bd95762c9ffe129c59efb08a91ee1.

## Fixed this round
- CI: replace Node20-based GitHub Actions `checkout@v4`, `setup-python@v5`, `upload-artifact@v4` with Node24-compatible `checkout@v5`, `setup-python@v6`, `upload-artifact@v6`. Keep Windows 2025, Python 3.12.10 x64, pinned dependencies and 10 packaged EXE smoke loops unchanged. GitHub removed Node20 from hosted runners on 2026-09-23; Node24 requires hosted runner >=2.327.1.
- Explicit network/SMB errors now have a separate Korean failure diagnosis, including in the existing file-analysis error dialog; ordinary permission, local EIO and full disk errors keep their own diagnosis. Advisory includes checking the output directory before retry if server response vanished after publishing.
- Scanning a source with a network disconnection must fail the whole scan. Previously `file_fingerprint()` failing to read bytes was swallowed as an optional fingerprint, which risked accepting a partial/unverified network directory. Permission/DLP read restriction continues to permit name-based grouping and merge-time error advice.
- Synthetic tests inject read, final ZIP write, and uncertain publish failures. They verify existing XLSX bytes untouched, scratch cleanup, no unverified scan, continuing healthy groups and no blind retry.

## Scope and evidence
- Before code: 102 baseline tests pass; targeted new tests failed in expected network/CI and GUI-scan cases.
- After code: 111 local unittest tests pass under Linux Xvfb; Python compileall successful.
- Network faults are synthetic errno/WinError injections, not live corporate SMB disconnects or DRM documents.
- At publication, verify exact GitHub commit, Windows Actions full suite, PyInstaller EXE smoke, and release assets. No claim before evidence.

## Further priorities
1. **P1 — corporate environment:** approved Windows 11 company PC, actual SMB share with authorized DRM samples, antivirus/EDR and VPN variations, and 8-hour repeated workload test. The hosted runner cannot exercise these conditions.
2. **P2 — remote write acknowledgement:** if an SMB server commits the final file before dropping the response, success is genuinely uncertain. UI now directs inspection; never silently auto-retry.
3. **P2 — sampling limitation:** source fingerprint samples 17 bounded windows on large files. Mutations outside samples can still escape detection. Full-file hashing increases SMB reads; requires performance/accuracy tradeoff decision.
4. **P2 — full Excel compatibility:** complex embedded objects and Excel formula recalc need representative Excel-based acceptance testing.
5. **P3 — long-run GUI logs and packaging scratch:** long log retention grows memory; PyInstaller onefile temporary extraction may be left after force kill. Avoid silently deleting user or unrelated process files.

## Rollback
Previous v1.3.18 Release remains available. Normal CSV/Excel merge, output filename suffixes, selected group checks, Korean GUI layout, and CSV text preservation remain unchanged.
