"""Real Qt workflow on isolated synthetic inputs; never uses production data paths."""
from pathlib import Path
import argparse
import sys
import tempfile
from unittest.mock import patch

PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--geometry", help="Logical client size, e.g. 840x580")
    options = parser.parse_args()
    from openpyxl import Workbook
    from PySide6.QtCore import QTimer
    from PySide6.QtWidgets import QApplication
    from e_help.qt_ui.application import ApplicationWindow
    from e_help.qt_ui.backend import Backend
    from e_help.qt_ui.theme import configure_application
    from e_help import history
    temporary_root = PROJECT.parent / "本地产物" / "E-help" / "build" / "qt-preview-runtime"
    temporary_root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ehelper-qt-preview-", dir=temporary_root) as temporary:
        root = Path(temporary)
        for name, rows in (("October-A.xlsx", [("示例客户A", 1280, "2026-10-01"), ("示例客户B", 890, "2026-10-01")]),
                           ("October-B.xlsx", [("示例客户A", .5, "2026-10-01")])):
            book = Workbook()
            book.active.append(["客户", "金额", "日期"])
            for row in rows:
                book.active.append(row)
            book.save(root / name)
            book.close()
        with patch.object(history, "get_history_path", return_value=str(root / "history.json")):
            app = QApplication(sys.argv[:1])
            configure_application(app)
            window = ApplicationWindow(Backend(root, root / "config", root / "logs"))
            window.setWindowTitle("E-help · Qt 业务验收（隔离合成数据）")
            if options.geometry:
                width, height = map(int, options.geometry.lower().split("x"))
                window.resize(width, height)
                window.setWindowTitle(f"E-help · Qt 窄窗口验收（{width}x{height}）")
            window.select_mode("monthly")
            window.selected_files = [str(root / name) for name in ("October-A.xlsx", "October-B.xlsx")]
            window.month.setCurrentText("2026-10")
            window.refresh_files()
            window.show()
            QTimer.singleShot(0, window.preview)
            return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
