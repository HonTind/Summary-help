"""Developer UI preview with only synthetic input and a disposable local data root.

Never run start_gui.py for screenshot QA against real user data. This entry does
not write screenshots, build products, contacts, repository source or user data.
"""
import argparse
import json
from pathlib import Path
import sys
import tempfile
import tkinter as tk
from unittest.mock import patch

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
from Excel自动汇总助手 import desktop_ui, utils, monthly, analysis
from Excel自动汇总助手.display import enable_high_dpi, configure_display


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--sample", action="store_true", help="Display explicitly synthetic monthly results")
    parser.add_argument("--geometry", default="1280x940")
    parser.add_argument("--density", type=float, help="Simulate startup layout density; not an OS setting")
    parser.add_argument("--metrics", action="store_true", help="Print only display and font measurements")
    options = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="ehelper-ui-preview-") as directory:
        sandbox = Path(directory)
        with patch.object(utils, "get_base_dir", return_value=directory), \
                patch.object(desktop_ui.gui, "get_base_dir", return_value=directory):
            root = tk.Tk()
            if options.density is not None:
                root.tk.call("tk", "scaling", options.density * 96 / 72)
                configure_display(root, ratio=options.density)
            app = desktop_ui.UnifiedApp(root)
            root.title("E助手 · 离线UI预览（临时示例）")
            root.geometry(options.geometry)
            if options.sample:
                from openpyxl import Workbook
                first, second = sandbox / "October-A.xlsx", sandbox / "October-B.xlsx"
                for path, rows in ((first, (("示例客户A", "1280.00"), ("示例客户B", "890.00"))),
                                   (second, (("示例客户A", "0.50"),))):
                    book = Workbook()
                    book.active.append(("客户", "金额"))
                    for row in rows:
                        book.active.append(row)
                    book.save(path)
                    book.close()
                window = app.workbench
                window.monthly_input.set_files([first, second])
                window.month.set("2026-10")
                window.template_name.set("标准月度规则")
                result = monthly.run_monthly([first, second], monthly.make_template(
                    "标准月度规则", analysis.SheetRules()), "2026-10")
                window._show_result(result)
                window.status.set("临时示例数据，仅用于界面验收；没有读取真实账单。")
            if options.metrics:
                from tkinter import font as tkfont
                root.update_idletasks()
                print(json.dumps({"density": root._display_density, "tk_scaling": root.tk.call("tk", "scaling"),
                                  "viewport": [root.winfo_width(), root.winfo_height()],
                                  "font": tkfont.Font(root=root, font=(root._display_font_family, 12)).actual()},
                                 ensure_ascii=False), flush=True)
            root.mainloop()


if __name__ == "__main__":
    enable_high_dpi()
    main()
