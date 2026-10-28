# Stability-only hardening design (approved audit follow-up)

Scope: user approved remediation of all prioritized findings in `ExcelSheetMerger_v1.3.12_안정성_정밀감사_20261028.md`, explicitly without adding/removing end-user features.

## Guarantees
- Inputs, filename grouping, time sheet names, checkboxes and UI layout, CSV string semantics, XLS/XLSX/XLSM styles/formulas/merges, EXE distribution remain identical for unchanged valid inputs.
- Failed folder scan cannot use results from a prior folder; successful scans remain unchanged.
- For sparse XLSX iterate only instantiated cells; for mixed CSV/Excel write CSV sheets through a streaming workbook and recombine their sheet XML into a normal workbook without materializing large CSV cells.
- Strongly isolated per-merge staging directories with a held process-level OS lock, safe recovery only after 24 hours and only when lock can be acquired and marker verified. Never delete unknown/legacy scratch files.
- Work record is written atomically; failure to write the log is reported separately from successful workbook creation.
- Detect input changes between scan and merge and during merge using file identity, length and mtime. Do not change CSV encoding policy.
- Bound GUI event-processing batches to keep interaction responsive, preserving queue ordering and status.
- Pin runtime build dependencies and Windows runner; add post-PyInstaller smoke test that runs the actual executable on sample CSV and XLSX inputs without opening the GUI, does not affect normal usage.

Alternatives rejected: treating workbook data as strings for memory optimization (breaks types); unguarded temp deletion (could delete a live process file); loading entire XML with `ZipFile.read` (memory spike); mutable scan report with cross-folder coupling; a normal save directly to shared output path (partial/overwritten file).

## Conditions
- `SourceFileError` remains a subclass of ValueError with original filename and cause; no DRM bypass.
- All existing public functions and UI controls stay intact for normal users.
- Release only with full successful Windows tests plus EXE smoke, preserving v1.3.12.

## Acceptance
- 67 baseline tests pass plus focused regressions and memory workload.
- A success → B failed scan → no A output in B and disabled merge.
- 700×700 sparse styled Excel does not allocate a full 700×700 cell grid.
- 60k×8 mixed input runs with bounded memory and equivalent XLSX workbook cells/styles/merges.
- Concurrent merge jobs do not remove each other's temp dirs and generate unique final names.
- Crash residue cleanup respects locks, marker and retention age.
- Missing/changed input after scan cannot produce an inconsistent result.
- Failed record write is surfaced as a warning but does not misreport data creation as failed.
- Windows CI launches the actual packaged EXE and checks generated XLSX and exit status.
