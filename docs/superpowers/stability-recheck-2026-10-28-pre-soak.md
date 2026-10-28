# Excel Sheet Merger — pre-soak reliability review

**Date:** 2026-10-28. **Baseline:** GitHub v1.3.22, main commit `b5e8287059b6cac4cd7ea5f007a8d6e8ba77a86d`, 122 local unit tests.

## Implemented before long-run testing

1. **Windows reparse ancestry detection:** before using full 8–32MiB fixed-drive fingerprinting, walk each source component by `lstat()` from root to leaf. An ancestor symlink, junction/reparse attribute, or unknown metadata makes the full-hash test ineligible; the existing 17 × 64KiB bounded sampling remains. Previous code checked only the terminal file, allowing a C-drive junction to SMB to be misclassified. Reading reparse metadata does not resolve the target.
2. **Result-folder callback error:** `_open_results()` now catches `OSError` from `mkdir()` and Explorer opening and uses the existing localized diagnosis. Previously these errors escaped to Tkinter callback handling. Successful behavior and layout remain identical.
3. **Tests added:** actual ancestor symlink; simulated reparse attribute; inaccessible component; optional actual Windows junction; `mkdir()` permission failure; optional Windows `os.startfile` failure. Both fixes were tested RED→GREEN.

## Regression and exclusions

- Full GUI/unittest suite run after modifications and Windows Actions verification must be reported against the exact released commit; no claim of Windows success until the workflow finishes.
- CSV/XLSX/XLSM/XLS decoding, group/date naming, formula/style transfer, group checkboxes, duplicate-output numbering, and existing log history remain unchanged.
- **Known limitations remain:** 17-window sampling is not a complete integrity guarantee for SMB or >32MiB sources; DRM/network share/antivirus real-world behavior; advanced workbook charts/pivot/external links; and full-history UI log memory growth. Replacing all visible logs with a rolling buffer would change behavior and needs distinct UX acceptance.
- **Final stage (after remaining code fixes):** Codex on an authorized Windows 11 PC runs an 8-hour repeated mixed CSV/Excel soak and fault-injection scenario with sanitized inputs; track process private bytes, file handles, CPU, disk, output CRC, stable hashes, log integrity, and network outage recovery. Never run this earlier and never claim it has already happened.
