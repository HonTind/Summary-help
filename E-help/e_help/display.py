"""Startup display policy and logical-pixel metrics; no OS settings are changed."""
from __future__ import annotations

import math
import sys
from tkinter import font as tkfont, ttk


def enable_high_dpi():
    """Opt this process into system DPI awareness before creating any Tk window.

    Tk 8.6 does not guarantee dynamic rescaling of existing widgets. Use startup
    system density, and preserve awareness already selected by an EXE manifest.
    """
    if sys.platform != "win32":
        return "not-windows"
    import ctypes
    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        setter = user32.SetProcessDpiAwarenessContext
        setter.argtypes = [ctypes.c_void_p]
        setter.restype = ctypes.c_bool
        if setter(ctypes.c_void_p(-2)):  # DPI_AWARENESS_CONTEXT_SYSTEM_AWARE
            return "system-aware"
        if ctypes.get_last_error() == 5:  # Already configured: don't override it.
            return "already-configured"
        return "unavailable"
    except (AttributeError, OSError):
        try:
            return "system-aware" if ctypes.windll.user32.SetProcessDPIAware() else "unavailable"
        except (AttributeError, OSError):
            return "unavailable"


def density(widget):
    root = getattr(widget, "root", None) or widget.winfo_toplevel()
    return getattr(root, "_display_density", 1.0)


def px(widget, value):
    if isinstance(value, (tuple, list)):
        return tuple(px(widget, item) for item in value)
    return round(float(value) * density(widget))


def configure_display(root, *, ratio=None):
    """Set density before creating widgets; point fonts must not be scaled twice."""
    ratio = root.winfo_fpixels("1i") / 96 if ratio is None else float(ratio)
    if not math.isfinite(ratio) or not 0.75 <= ratio <= 4:
        raise ValueError("Display density must be between 0.75 and 4.")
    root._display_density = ratio
    families = set(tkfont.families(root))
    family = next((name for name in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI")
                   if name in families), "TkDefaultFont")
    root._display_font_family = family
    return ratio


def window_bounds(root):
    """Fit the startup client area, including on small high-density displays.

    Reserve room for window chrome and the taskbar. The scrollable settings
    panel handles a client area smaller than the usual logical minimum.
    """
    available = (max(1, root.winfo_screenwidth() - px(root, 80)),
                 max(1, root.winfo_screenheight() - px(root, 120)))
    initial = tuple(min(px(root, value), limit) for value, limit in zip((1168, 828), available))
    minimum = tuple(min(px(root, value), limit) for value, limit in zip((840, 580), available))
    return initial, minimum


def configure_styles(root):
    """Use named flat controls while retaining the application's native theme."""
    style = ttk.Style(root)
    body = (root._display_font_family, 12)
    secondary = "#667085"
    # Startup awareness also affects point fonts in existing native dialogs.
    # Preserve their fonts/theme; only keep the default tree rows tall enough.
    if not hasattr(root, "_legacy_tree_rowheight"):
        root._legacy_tree_rowheight = int(style.lookup("Treeview", "rowheight") or 20)
    legacy_font = style.lookup("Treeview", "font") or "TkDefaultFont"
    legacy_line = tkfont.Font(root=root, font=legacy_font).metrics("linespace")
    style.configure("Treeview", rowheight=max(root._legacy_tree_rowheight, legacy_line + px(root, 4)))
    for name, original in (("Desktop.entry.field", "Entry.field"),
                           ("Desktop.combo.field", "Combobox.field"),
                           ("Desktop.combo.downarrow", "Combobox.downarrow"),
                           ("Desktop.notebook.client", "Notebook.client"),
                           ("Desktop.tree.field", "Treeview.field"),
                           ("Desktop.heading.cell", "Treeheading.cell"),
                           ("Desktop.heading.border", "Treeheading.border"),
                           ("Desktop.pager.border", "Button.border")):
        if name not in style.element_names():
            style.element_create(name, "from", "clam", original)
    style.layout("Desktop.TEntry", [("Desktop.entry.field", {"sticky": "nswe", "children": [
        ("Entry.padding", {"sticky": "nswe", "children": [("Entry.textarea", {"sticky": "nswe"})]})]})])
    style.layout("Desktop.TCombobox", [("Desktop.combo.field", {"sticky": "nswe", "children": [
        ("Desktop.combo.downarrow", {"side": "right", "sticky": "ns"}),
        ("Combobox.padding", {"sticky": "nswe", "children": [("Combobox.textarea", {"sticky": "nswe"})]})]})])
    for name in ("Desktop.TEntry", "Desktop.TCombobox"):
        style.configure(name, fieldbackground="white", foreground="#1D1D1F", bordercolor="#D8DEE8",
                        lightcolor="white", darkcolor="white", padding=px(root, (8, 7)), font=body,
                        arrowsize=px(root, 15))
        style.map(name, bordercolor=[("focus", "#0071E3")],
                  fieldbackground=[("disabled", "#F4F6F9"), ("readonly", "white")],
                  foreground=[("disabled", "#98A2B3")])
    style.layout("Desktop.TNotebook", [("Desktop.notebook.client", {"sticky": "nswe"})])
    style.configure("Desktop.TNotebook", background="white", bordercolor="white", lightcolor="white",
                    darkcolor="white", borderwidth=0, tabmargins=(0, 0, 0, 0))
    style.layout("Desktop.TNotebook.Tab", [("Notebook.padding", {"sticky": "nswe", "children": [
        ("Notebook.label", {"sticky": "nswe"})]})])
    style.configure("Desktop.TNotebook.Tab", font=body, foreground=secondary,
                    background="white", padding=px(root, (14, 10)))
    style.map("Desktop.TNotebook.Tab", foreground=[("selected", "#0071E3")],
              background=[("selected", "white"), ("active", "#F5F8FD")])
    style.configure("Desktop.TCheckbutton", font=body, background="white", foreground="#344054")
    style.map("Desktop.TCheckbutton", background=[("active", "white")],
              foreground=[("disabled", "#98A2B3")])
    style.layout("Desktop.Pager.TButton", [("Desktop.pager.border", {"sticky": "nswe", "children": [
        ("Button.focus", {"sticky": "nswe", "children": [("Button.padding", {"sticky": "nswe", "children": [
            ("Button.label", {"sticky": "nswe"})]})]})]})])
    style.configure("Desktop.Pager.TButton", font=(body[0], 10), padding=px(root, (6, 4)),
                    background="white", foreground="#344054", bordercolor="#D8DEE8",
                    lightcolor="white", darkcolor="white")
    style.map("Desktop.Pager.TButton", foreground=[("disabled", "#98A2B3")],
              background=[("active", "#EFF6FF"), ("disabled", "#F8FAFC")])
    style.configure("Desktop.Pager.TLabel", background="white", foreground=secondary, font=(body[0], 10))
    style.layout("Desktop.Treeview", [("Desktop.tree.field", {"sticky": "nswe", "children": [
        ("Treeview.padding", {"sticky": "nswe", "children": [("Treeview.treearea", {"sticky": "nswe"})]})]})])
    style.layout("Desktop.Treeview.Heading", [("Desktop.heading.cell", {"sticky": "nswe"}),
        ("Desktop.heading.border", {"sticky": "nswe", "children": [("Treeheading.padding", {
            "sticky": "nswe", "children": [("Treeheading.image", {"side": "right", "sticky": ""}),
                                             ("Treeheading.text", {"sticky": "we"})]})]})])
    line = tkfont.Font(root=root, font=body).metrics("linespace")
    style.configure("Desktop.Treeview", font=body, rowheight=max(px(root, 38), line + px(root, 12)),
                    background="white", fieldbackground="white", foreground="#344054",
                    bordercolor="#E1E6EF", lightcolor="#E1E6EF", darkcolor="#E1E6EF", borderwidth=1)
    style.configure("Desktop.Treeview.Heading", font=(body[0], 12, "bold"),
                    background="#EEF1F6", foreground="#1D1D1F", bordercolor="#E1E6EF",
                    lightcolor="#EEF1F6", darkcolor="#EEF1F6", padding=px(root, (12, 10)))
    style.map("Desktop.Treeview", background=[("selected", "#E8F1FE")],
              foreground=[("selected", "#1D1D1F")])
    style.configure("Desktop.FileList.Treeview", font=body, rowheight=px(root, 32),
                    background="#F8FAFC", fieldbackground="#F8FAFC", bordercolor="#E4E8F0",
                    lightcolor="#E4E8F0", darkcolor="#E4E8F0")
    style.map("Desktop.FileList.Treeview", foreground=[("disabled", "#98A2B3"), ("selected", "#344054")],
              background=[("selected", "#E8F1FE")])
    return style
