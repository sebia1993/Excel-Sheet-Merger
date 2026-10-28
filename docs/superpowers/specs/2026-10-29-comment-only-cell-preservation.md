# Blank Comment-Only Cell Preservation (2026-10-29)

## Purpose
Keep existing Excel Sheet Merger behavior unchanged except for lossless preservation of notes attached to empty Excel source cells. Input formats, group checkbox UI, time-named sheets, original preservation and result version numbering remain unchanged. Eight-hour Windows soak is deferred.

## Reproduced defect
A real source XLSX contains A1 text and an unstyled blank D20 with a comment. openpyxl loads D20 with value=None, has_style=False, and comment not None. _append_openpyxl previously skipped it before copying note metadata, silently losing the comment.

## Selected design
- Change the early skip predicate in _append_openpyxl to skip only when value is None and no style, comment or hyperlink exists.
- Keep the existing check that a source worksheet must contain an actual nonempty data value. A comment-only workbook does not become a data-filled workbook.
- Keep the same output and original files, unchanged CSV text rules, formula/style handling and no network access.
- Real XLSX end-to-end and direct _append_openpyxl regressions verify note text/author, empty cell value, sheet names, CSV ID 00123, numbered output and original bytes.
- No expensive full blank-cell materialization. No GUI or log behavior changes.

## Alternatives and tradeoffs
- Leaving skip unchanged loses user notes; rejected.
- Always copying every blank cell costs unbounded sparse worksheet memory; rejected.
- Metadata-aware skip is minimal and preserves existing performance for ordinary empty cells.

## GUI log follow-up
30k fully displayed lines increased RSS by around 7MiB on Linux/Xvfb. Current contract and regression tests require complete in-app history. Do not prune or truncate it in this stability change. Any disk-backed virtualization needs separate design and equivalence testing.
