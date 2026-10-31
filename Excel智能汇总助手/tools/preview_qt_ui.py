"""Standalone offline Qt preview; never launches or reads the production app."""
import argparse
import json
from pathlib import Path
import sys

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--geometry", help="Logical client WxH, e.g. 1168x824")
    parser.add_argument("--metrics", action="store_true")
    options = parser.parse_args()
    from PySide6.QtWidgets import QApplication
    from PySide6.QtGui import QFontInfo
    from Excel自动汇总助手.qt_ui.theme import configure_application
    from Excel自动汇总助手.qt_ui.window import PreviewWindow
    app = QApplication(sys.argv[:1])
    app.setApplicationName("E助手 Qt界面预览")
    configure_application(app)
    window = PreviewWindow()
    if options.geometry:
        width, height = map(int, options.geometry.lower().split("x"))
        window.resize(width, height)
    window.show()
    app.processEvents()
    if options.metrics:
        print(json.dumps({"logical_client": [window.width(), window.height()],
            "device_pixel_ratio": window.devicePixelRatioF(), "font_family": QFontInfo(app.font()).family(),
            "demo_only": True}, ensure_ascii=False), flush=True)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
