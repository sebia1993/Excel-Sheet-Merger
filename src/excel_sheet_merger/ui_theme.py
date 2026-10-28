"""Minimal Office UI palette and ttkbootstrap-compatible style tokens.

Keep third-party theme initialization at runtime: importing the CSV engine never
requires a GUI or a display. On Linux without ttkbootstrap, clam is available
for offline tests; Windows production builds package ttkbootstrap explicitly.
"""

from __future__ import annotations

import os
from tkinter import ttk

COLORS = {
    "background": "#F6F8F7",
    "surface": "#FFFFFF",
    "surface_subtle": "#EEF3F0",
    "primary": "#166D53",
    "primary_hover": "#116149",
    "text": "#1F2937",
    "muted": "#6B7280",
    "border": "#DCE5DF",
    "success": "#13795B",
    "warning": "#946011",
    "error": "#BA3A34",
    "white": "#FFFFFF",
    "disabled": "#CBD9D0",
}

FONT = "Malgun Gothic"


def configure_theme(root) -> ttk.Style:
    """Apply shared styles while preserving native Tk/ttk widget APIs.

    Windows is strict: a packaging error that omits ttkbootstrap must be
    discovered at startup, not silently degrade a released EXE.
    """
    try:
        from ttkbootstrap import Style as BootstrapStyle
    except ImportError:
        if os.name == "nt":
            raise
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
    else:
        style = BootstrapStyle(theme="litera")

    c = COLORS
    style.configure("Office.Page.TFrame", background=c["background"])
    style.configure("Office.Card.TFrame", background=c["surface"])
    style.configure("Office.Subtle.TFrame", background=c["surface_subtle"])
    style.configure("Office.Title.TLabel", background=c["background"], foreground=c["text"], font=(FONT, 18, "bold"))
    style.configure("Office.Subtitle.TLabel", background=c["background"], foreground=c["muted"], font=(FONT, 9))
    style.configure("Office.Eyebrow.TLabel", background=c["surface"], foreground=c["muted"], font=(FONT, 9))
    style.configure("Office.Heading.TLabel", background=c["surface"], foreground=c["text"], font=(FONT, 11, "bold"))
    style.configure("Office.Body.TLabel", background=c["surface"], foreground=c["text"], font=(FONT, 9))
    style.configure("Office.Hint.TLabel", background=c["surface"], foreground=c["muted"], font=(FONT, 9))
    style.configure("Office.Subtle.TLabel", background=c["surface_subtle"], foreground=c["muted"], font=(FONT, 9))
    style.configure("Office.Metric.TLabel", background=c["surface"], foreground=c["text"], font=(FONT, 17, "bold"))
    style.configure("Office.MetricHint.TLabel", background=c["surface"], foreground=c["muted"], font=(FONT, 9))
    style.configure("Office.Pill.TLabel", background=c["surface_subtle"], foreground=c["primary"], font=(FONT, 9, "bold"), padding=(10, 5))
    style.configure("Office.Status.TLabel", background=c["surface"], foreground=c["text"], font=(FONT, 9))
    style.configure("Office.Failure.TLabel", background=c["surface_subtle"],
                    foreground=c["error"], font=(FONT, 9), padding=(6, 3))

    style.configure("Office.Primary.TButton", background=c["primary"], foreground=c["white"], font=(FONT, 10, "bold"), padding=(15, 7), relief="flat", borderwidth=0)
    style.map("Office.Primary.TButton",
              background=[("disabled", c["disabled"]), ("pressed", c["primary_hover"]), ("active", c["primary_hover"])],
              foreground=[("disabled", c["muted"]), ("!disabled", c["white"])])
    style.configure("Office.Secondary.TButton", background=c["surface_subtle"], foreground=c["text"], font=(FONT, 9, "bold"), padding=(12, 7), relief="flat", borderwidth=1)
    style.map("Office.Secondary.TButton", background=[("disabled", c["background"]), ("active", c["border"])])
    style.configure("Office.Link.TButton", background=c["surface"], foreground=c["muted"], font=(FONT, 9), padding=(7, 2), relief="flat", borderwidth=0)
    style.map("Office.Link.TButton", foreground=[("active", c["primary"])])
    style.configure("Office.Path.TEntry", padding=7, font=(FONT, 10))
    style.map("Office.Path.TEntry", fieldbackground=[("readonly", c["background"])], foreground=[("readonly", c["text"])])

    style.configure("Office.Treeview", background=c["surface"], fieldbackground=c["surface"], foreground=c["text"], borderwidth=0, rowheight=32, font=(FONT, 9))
    style.configure("Office.Treeview.Heading", background=c["surface_subtle"], foreground=c["text"], font=(FONT, 9, "bold"), padding=(8, 9), relief="flat")
    style.map("Office.Treeview", background=[("selected", "#DBEDE5")], foreground=[("selected", c["text"])])
    style.configure("Office.Horizontal.TProgressbar", troughcolor=c["surface_subtle"], background=c["primary"], borderwidth=0, thickness=8)
    # ttkbootstrap 2.2.6's scrollbar thumb uses a Tk field element whose
    # focuscolor must be valid even with focuswidth=0. Without it the first
    # render raises a Tcl background error outside normal Python callbacks.
    # A color-only default preserves the existing thumb geometry/appearance.
    style.configure("Vertical.TScrollbar", focuscolor=c["primary"])
    # Bind to the actual root, not global Tk.destroy. Toplevel bindings also
    # receive child-destroy events: only release at this root's own teardown.
    root.bind("<Destroy>",
              lambda event: release_theme(root) if event.widget is root else None,
              add="+")
    return style


def release_theme(root) -> None:
    """Release only the Bootstrap caches owned by this root's interpreter.

    MergeApp subclasses tk.Tk rather than ttkbootstrap.App, so it must perform
    the same Style/Fonts teardown as the pinned App implementation. Never
    invalidate a theme belonging to another still-live root.
    """
    style = getattr(root, "theme_style", None)
    if style is None or getattr(style, "tk", None) is not root.tk:
        return
    try:
        from ttkbootstrap import Style as BootstrapStyle
        from ttkbootstrap.utils.fonts import Fonts
    except ImportError:
        return  # Native ttk fallback on offline non-Windows development hosts.
    if BootstrapStyle.instance is style:
        BootstrapStyle.instance = None
        Fonts.reset()
