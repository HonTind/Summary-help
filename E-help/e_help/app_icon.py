"""Offline application icons, independent of writable user-data paths."""
from pathlib import Path
import sys
import tkinter as tk


def apply_app_icon(root):
    """Set this root's icon and the default for child windows; fail gracefully.

    PyInstaller preserves the package's __file__ beneath its bundle directory,
    so this also resolves assets in both onedir and onefile bundles.
    """
    assets = Path(__file__).resolve().parent.parent / "assets"
    if sys.platform == "win32":
        try:
            root.iconbitmap(default=str(assets / "app-icon.ico"))
            return True
        except (tk.TclError, OSError):
            pass
    try:
        image = tk.PhotoImage(master=root, file=str(assets / "app-icon.png"))
        root.iconphoto(True, image)
        root._app_icon_image = image
        return True
    except (tk.TclError, OSError):
        return False
