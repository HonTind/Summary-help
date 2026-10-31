"""Display regressions: isolated data, simulated density, no OS setting changes."""
import base64
from contextlib import contextmanager, ExitStack
import ctypes
import importlib
from pathlib import Path
import socket
import struct
import sys
import tempfile
import tkinter as tk
from tkinter import font as tkfont, ttk
from types import SimpleNamespace
import unittest
import urllib.request
import webbrowser
from unittest.mock import Mock, patch

PROJECT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(PROJECT))
display = importlib.import_module("e_help.display")
ui = importlib.import_module("e_help.desktop_ui")
assets = importlib.import_module("e_help.ui_assets")
utils = importlib.import_module("e_help.utils")


class DisplayPolicyTests(unittest.TestCase):
    def test_non_windows_does_not_load_windows_libraries(self):
        with patch.object(display.sys, "platform", "linux"), patch.object(ctypes, "WinDLL", create=True) as loader:
            self.assertEqual(display.enable_high_dpi(), "not-windows")
            loader.assert_not_called()

    def test_startup_awareness_and_existing_manifest_are_preserved(self):
        for accepted, error, expected in ((True, 0, "system-aware"),
                                           (False, 5, "already-configured"),
                                           (False, 87, "unavailable")):
            with self.subTest(expected=expected):
                setter = Mock(return_value=accepted)
                library = SimpleNamespace(SetProcessDpiAwarenessContext=setter)
                with patch.object(display.sys, "platform", "win32"), \
                        patch.object(ctypes, "WinDLL", return_value=library, create=True), \
                        patch.object(ctypes, "get_last_error", return_value=error, create=True):
                    self.assertEqual(display.enable_high_dpi(), expected)
                self.assertEqual(setter.call_args.args[0].value, ctypes.c_void_p(-2).value)
                setter.assert_called_once()

    def test_older_windows_api_fallback(self):
        old = Mock(return_value=True)
        with patch.object(display.sys, "platform", "win32"), \
                patch.object(ctypes, "WinDLL", side_effect=AttributeError, create=True), \
                patch.object(ctypes, "windll", SimpleNamespace(user32=SimpleNamespace(SetProcessDPIAware=old)), create=True):
            self.assertEqual(display.enable_high_dpi(), "system-aware")
            old.assert_called_once()

    def test_window_minimum_fits_small_high_density_screen(self):
        root = SimpleNamespace(_display_density=1.25, winfo_screenwidth=lambda: 1366,
                               winfo_screenheight=lambda: 768)
        root.winfo_toplevel = lambda: root
        initial, minimum = display.window_bounds(root)
        self.assertEqual(initial, (1266, 618))
        self.assertEqual(minimum, (1050, 618))
        self.assertEqual(display.px(root, (8, 12)), (10, 15))


class IconAssetsTests(unittest.TestCase):
    def test_icons_have_native_density_variants_and_alpha(self):
        for ratio in (1, 1.25, 1.5, 2, 3, 4):
            size = round(24 * ratio)
            self.assertIn(size, assets.ICON_PNGS)
            for state in ("normal", "active"):
                for name, encoded in assets.ICON_PNGS[size][state].items():
                    data = base64.b64decode(encoded)
                    self.assertEqual(data[:8], b"\x89PNG\r\n\x1a\n")
                    self.assertEqual(struct.unpack(">II", data[16:24]), (size, size), name)
                    self.assertEqual(data[25], 6, name)  # RGBA: no opaque crop background.
            self.assertNotEqual(assets.ICON_PNGS[size]["normal"]["workbench"],
                                assets.ICON_PNGS[size]["active"]["workbench"])

    def test_original_brand_is_present_at_each_startup_density(self):
        for ratio in (1, 1.25, 1.5, 2):
            size = round(32 * ratio)
            data = base64.b64decode(assets.BRAND_PNGS[size])
            self.assertEqual(struct.unpack(">II", data[16:24]), (size, size))
        self.assertTrue((PROJECT / "assets/ui/FEATHER-LICENSE.txt").is_file())


class DisplayLayoutTests(unittest.TestCase):
    @contextmanager
    def app(self, ratio, screen=None):
        with ExitStack() as stack:
            directory = stack.enter_context(tempfile.TemporaryDirectory(prefix="ehelper-display-data-"))
            for target in (utils, ui.gui):
                stack.enter_context(patch.object(target, "get_base_dir", return_value=directory))
            def no_network(*_args, **_kwargs):
                raise AssertionError("Display initialization attempted network access")
            for target, name in ((socket.socket, "connect"), (socket, "create_connection"),
                                 (urllib.request, "urlopen"), (webbrowser, "open")):
                stack.enter_context(patch.object(target, name, side_effect=no_network))
            root = tk.Tk()
            root.withdraw()
            root.tk.call("tk", "scaling", ratio * 96 / 72)
            original_fonts = {name: tkfont.nametofont(name, root=root).actual()
                              for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont")}
            display.configure_display(root, ratio=ratio)
            if screen:
                stack.enter_context(patch.object(root, "winfo_screenwidth", return_value=screen[0]))
                stack.enter_context(patch.object(root, "winfo_screenheight", return_value=screen[1]))
            errors = []
            root.report_callback_exception = lambda *error: errors.append(error)
            try:
                app = ui.UnifiedApp(root)
                root.deiconify()
                root.update()
                yield root, app, original_fonts
            finally:
                root.destroy()
                self.assertFalse(errors, errors)

    def inside(self, root, control):
        self.assertTrue(control.winfo_ismapped())
        x = control.winfo_rootx() - root.winfo_rootx()
        y = control.winfo_rooty() - root.winfo_rooty()
        self.assertGreaterEqual(x, 0)
        self.assertGreaterEqual(y, 0)
        self.assertLessEqual(x + control.winfo_width(), root.winfo_width())
        self.assertLessEqual(y + control.winfo_height(), root.winfo_height())

    def test_four_startup_densities_keep_fonts_icons_and_actions_reachable(self):
        for ratio in (1, 1.25, 1.5, 2):
            with self.subTest(ratio=ratio), self.app(ratio) as (root, app, original_fonts):
                for name, original in original_fonts.items():
                    self.assertEqual(tkfont.nametofont(name, root=root).actual(), original)
                self.assertEqual(app.icons["workbench"].width(), round(24 * ratio))
                self.assertEqual(app.icons["brand"].width(), round(32 * ratio))
                root.geometry(f"{display.px(root, 840)}x{display.px(root, 580)}")
                root.update()
                for mode in ("monthly", "compare"):
                    app.select_mode(mode)
                    root.update()
                    for control in (app.workbench.preview_button, app.workbench.export_button,
                                    app.workbench.cancel_button, app.help_button):
                        self.inside(root, control)
                style = ttk.Style(root)
                line = tkfont.Font(root=root, font=style.lookup("Desktop.Treeview", "font")).metrics("linespace")
                self.assertGreater(int(style.lookup("Desktop.Treeview", "rowheight")), line)
                old_line = tkfont.Font(root=root, font=style.lookup("Treeview", "font") or "TkDefaultFont").metrics("linespace")
                self.assertGreater(int(style.lookup("Treeview", "rowheight")), old_line)
                old_height = int(style.lookup("Treeview", "rowheight"))
                display.configure_styles(root)
                self.assertEqual(int(style.lookup("Treeview", "rowheight")), old_height)
                app.select_mode("monthly")
                root.update()
                notebook = app.workbench.views
                for tab in range(3):
                    notebook.select(tab)
                    root.update()
                    marker = app.workbench.view_indicator
                    self.assertEqual(notebook.index(f"@{marker.winfo_x()},{display.px(root, 12)}"), tab)
                    self.assertEqual(notebook.index(f"@{marker.winfo_x() + marker.winfo_width() - 1},{display.px(root, 12)}"), tab)
                app.show_page("help")
                root.update()
                app.show_page("history")
                root.update()
                self.assertEqual(tkfont.Font(root=root, font=app.history_filter_box.cget("font")).actual("size"), 12)

    def test_small_125_percent_screen_keeps_footer_visible(self):
        with self.app(1.25, screen=(1366, 768)) as (root, app, _fonts):
            self.assertLessEqual(root.winfo_height(), 618)
            minimum = root.minsize()
            root.geometry(f"{minimum[0]}x{minimum[1]}")
            root.update()
            for control in (app.workbench.preview_button, app.workbench.export_button, app.help_button):
                self.inside(root, control)
            pager = app.workbench.summary
            self.assertEqual(str(pager.previous.cget("compound")), "none")
            self.assertEqual(pager.label.winfo_width(), pager.label.winfo_reqwidth())
            hint = app.workbench.empty_hint
            self.assertLessEqual(int(hint.cget("wraplength")), hint.winfo_width() - 2 * int(hint.cget("padx")))
            scroll = app.workbench.settings_scroll
            self.assertTrue(scroll.scrollbar.winfo_ismapped())
            scroll.scroll("moveto", "1")
            root.update()
            self.assertGreater(scroll.offset, 0)
            self.assertLessEqual(scroll.inner.winfo_y() + scroll.inner.winfo_height(),
                                 scroll.winfo_height() + int(scroll.cget("highlightthickness")))


if __name__ == "__main__":
    unittest.main()
