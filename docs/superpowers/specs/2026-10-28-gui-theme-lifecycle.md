# GUI theme lifecycle stability

Baseline: v1.3.27 / 681d145bcc8d305742a14c01594535d197480e1d.
The user authorizes internal stability fixes and deployment without routine approval;
no new features, file semantics, visible design changes, or eight-hour soak test.

## Reproduced defects
- With ttkbootstrap 2.2.6, the scrollbar thumb field element consumes `focuscolor`
  even though focuswidth is zero. The default Vertical.TScrollbar leaves that color
  empty, causing Tcl `unknown color name ""` on first rendering. Windows diagnostic
  run 40707317680 found one background error while all 150 unittest cases passed.
- MergeApp inherits plain tk.Tk, not ttkbootstrap.App; its destroy() does not release
  the library's process-wide Style singleton and named font cache. A later MergeApp
  receives a style tied to the already destroyed Tcl interpreter.

## Chosen repair
Set a valid existing-palette focus color for the vertical scrollbar (no width,
layout, or design change). Release only this root's owned Bootstrap singleton and
font wrappers on this root's own Destroy event, mirroring the pinned library's App.destroy contract.
Do not clear another live root's theme. Keep the non-Bootstrap Linux fallback.

Alternatives rejected: hiding Tcl errors; switching UI frameworks/themes; weakening
existing layout assertions; globally monkey-patching third-party theme internals.

## Acceptance
Fresh-process actual Tk rendering must report zero background/callback errors;
repeated root creation must bind style to each current interpreter; scrollbars,
preview rows, checkbox selection and keyboard focus must remain usable. Run all
existing tests, Windows CI and packaged EXE smoke. Keep the separate diagnostic
workflow off main. Report platform limitations and any failing tests honestly.

## Ruling after real-theme Windows validation
Fresh-root styling exposed an existing 900x640 failure-banner layout defect: the
preview was 60px in Windows (requirement >=70px); original Linux also failed.
Keep all failure text/logs and normal spacing. Only while failure guidance is
visible below 700px height, compact card padding to at most 3px and root vertical
padding to 5px. Restore original spacing on resizing or clearing failures. This
is a minimum-window correctness repair, not a new control or theme redesign.
