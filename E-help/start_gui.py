"""
E-help v1.5.4 — 程序唯一正式入口

启动方式:
    python start_gui.py
    或打包: pyinstaller start_gui.py
"""

import sys


def main():
    if "--legacy-ui" in sys.argv[1:]:
        from e_help.gui import main as launch
    else:
        from e_help.qt_ui.main import main as launch
    return launch()

if __name__ == "__main__":
    raise SystemExit(main())
