# Excel Source Cell Lifetime Stability Design

Date: 2026-10-28. Prior authorization: keep improving stability without individual approval; 8-hour soak last.

## Intent
Reduce peak working memory when copying regular `.xlsx`/`.xlsm` source worksheets into XLSX results, without changing sheet values, types, formatting, merged ranges, formulas, output naming, or UI behavior. Existing openpyxl normal-mode load creates a dictionary of source Cell objects; copying into destination duplicates those objects until the source workbook object is collected.

## Alternatives
1. **Selected:** process already-instantiated source cells one at a time and release each source dictionary entry immediately after copying into destination. The source Workbook is private and read-only on disk, and the source dictionary is not accessed after cell transfer; worksheet dimensions, merges, freeze panes, and view settings are independent of the cell dictionary. This reduces simultaneous live source/output cells and leaves the published XLSX data unchanged.
2. Switch to read-only workbook streaming: reduces memory further but silently loses or complicates merged cells, comments, hyperlinks, layout and styling; not acceptable for no-feature-change requirement.
3. Retain existing `values()` loop and call `gc.collect()` at workbook close: may reclaim after each source, but cannot reduce peak during a single large source transfer.

## Constraints
- Never alter the original workbook on disk; source dictionary mutations are restricted to a temporary in-memory loaded workbook.
- Continue copying **all** non-empty or formatted cells including formula/string types, comments, hyperlinks, six style parts, merged-border behavior, row/column dimensions and freeze panes.
- Keep source ZIP safety limits and all data publishing/rename/rollback semantics unchanged.
- Avoid new dependencies and any additional user-facing settings or UI changes.

## Acceptance
- A regression test must observe source live-cell count shrinking *during* transfer (not only when the source workbook is closed). Watch test fail before production change.
- Compare formatted, merged, formula and sparse source Excel output against hand-checked values.
- Compare 2 × (12,000 × 8 cells) synthetic workbook peak RSS in isolated Linux Python subprocess before and after; report exact measured values without claiming production Windows equivalence.
- Existing full unittest suite passes, Windows EXE built and executed in CI before new Release claim.
- Re-audit any remaining risks; keep 8-hour corporate Windows soak for the last phase.
