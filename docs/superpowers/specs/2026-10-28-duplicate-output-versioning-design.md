# Duplicate Output Versioning Design — Excel Sheet Merger

**Status:** approved via explicit instruction to implement; **date:** 2026-10-28.

## User outcome
When the generated XLSX name already exists in `병합결과`, preserve it and create another workbook by appending ` (2)`, ` (3)`, ... immediately before `.xlsx`. Tell the user in the GUI log and end-of-job result that a conflict was found and a new name was used. Re-running the same group must produce another complete file, never skip or overwrite previous work.

## Approach and alternatives
1. **Atomic unique-name publishing (selected):** build each workbook into an independent temporary XLSX, then publish it atomically to the lowest available name. On Windows rename into a new name; on POSIX hard-link the temporary output exclusively. If the name is claimed concurrently, retry the next suffix. Robust against simultaneous runs, never overwrites. The original CSV/Excel reading and group selection pipelines remain unchanged.
2. File-exists precheck and regular save (rejected): subject to race/overwriting and partial files.
3. Timestamp filenames (rejected): less recognizable than successive suffixes, makes comparison harder.

## Scope and interfaces
- `src/excel_sheet_merger/merge.py`: output collision handling in `merge_group()`; add `MergeStatus.RENAMED` and return the *actual generated path* through existing `MergeResult.path`. `MergeStatus.CREATED` remains for the first successful file. Keep existing defensive `EXISTS` enum member only for backward compatibility; no longer return it for a name collision.
- `src/excel_sheet_merger/gui.py`: treat `CREATED` and `RENAMED` as created workbooks. Present old + new names in progress log for renamed results; report count of renamed workbooks (third number of existing 3-value `done` event) instead of skipped. Keep UI processing threading and checkbox semantics intact.
- README/release notes: update collision policy and new status messaging.
- Only XLSX result paths change. Input filename classification, time-named sheets, CSV-only streaming, Excel styles/formula reading, files already in output folder, per-group error isolation, no external Excel or Python installation, and UI layout remain unchanged.

## Acceptance cases
1. First pass -> `A사업장_20261028.xlsx` with original source sheets.
2. Second, third pass -> `A사업장_20261028 (2).xlsx`, `A사업장_20261028 (3).xlsx`; old workbook hashes/bytes remain identical, each new file has same sheet data.
3. If `(2)` already exists while `(3)` is free, choose `(3)`.
4. Distinct input combinations (CSV only, CSV+XLSX) both work.
5. If writing/parsing fails, no new result and no stranded `_병합중_*` file; existing results unchanged.
6. Simultaneous calls cannot replace each other's output. Refuse to mask other OS I/O errors as filename collisions.
7. UI records include the original name and the actual new name, and treats renamed output as a successful creation; checkbox selection and full merge unaffected.
8. Windows test suite and PyInstaller build pass before versioned Release. Existing release remains available.

## Risks and checks
- Windows/NTFS collision semantics: `os.rename` to an existing target raises `FileExistsError`; never substitute `os.replace`.
- POSIX atomic link: `os.link()` fails with `FileExistsError` if target taken; temporary file removed in `finally`.
- The rename loop can fail on filesystem path length, permissions, disk space, or unsupported hardlinks. Surface genuine errors; do not silently skip.
