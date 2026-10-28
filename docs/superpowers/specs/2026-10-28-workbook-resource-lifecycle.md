# Workbook resource lifecycle hardening

Baseline: `0cc28849e5da547ff5bb68a2a00ff4f9e857c568` (v1.3.26).
The user authorized ongoing stability-only fixes and release updates; no routine approval gate. Eight-hour company-PC tests remain deferred.

## Observed problems
- A valid CSV followed by an empty CSV fails without an output, but leaves openpyxl worksheet XML files in the OS TEMP directory. Closing Workbook alone does not clean these files.
- A malformed XLSX raises a useful conversion error, but its input ZIP handle can remain open while the exception/traceback is retained.
- Normal and write-only worksheet serialization can fail after a temporary XML file has been created; the library's normal success-only cleanup then does not run.

## Approach
- Own the input binary stream through a context manager around normal-mode load_workbook; retain filename and exception chains.
- Add an isolated openpyxl 3.1.5 save adapter which preserves the existing worksheet serialization, manifest and drawings, but closes its ZIP on failures and cleans each writer in a finally block.
- Finalize write-only worksheet generators and clean only writers reachable from the current task's two workbooks, including when a later input fails before saving. Never scan TEMP or delete based on global temp-file lists.
- Keep original data/filenames/UI/classification/ZIP safety thresholds and source validation unchanged. Cleanup errors must not replace the primary conversion error; OS refusal to delete still cannot be guaranteed away.

## Alternatives not selected
Changing global tempfile paths, global writer monkeypatching or removing every openpyxl temporary file could affect other active work. Relaxing high-ratio ZIP limits solely for compatibility could reintroduce resource exhaustion; not part of this fix.

## Acceptance
No owned TEMP XML after ordinary handled input/save failures; unrelated live writers survive. Source/read and result/write handles are closed even when errors are retained. Original ENOSPC and conversion error reasons survive secondary cleanup errors. Successful CSV/XLSX and XLSM data, formulas, style and ZIP member contents remain unchanged. Full Windows tests and packaged EXE smoke must succeed before claiming a release.
