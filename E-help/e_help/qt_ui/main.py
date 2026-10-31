"""Production local Qt entry. Paths retain the existing source/EXE behavior."""
import os
import sys
from PySide6.QtWidgets import QApplication
from ..utils import get_base_dir, get_config_dir, get_logs_dir
from .application import ApplicationWindow
from .backend import Backend
from .theme import configure_application


def main():
    logs = get_logs_dir()
    log_path = os.path.join(logs, "startup.log")

    def mark(value):
        with open(log_path, "a", encoding="utf-8") as handle:
            handle.write(value + "\n")

    mark("start")
    app = QApplication(sys.argv[:1])
    app.setApplicationName("E-help")
    configure_application(app)
    # Keep the established launch-check log markers compatible.
    mark("create SummaryApp begin")
    window = ApplicationWindow(Backend(get_base_dir(), get_config_dir(), logs))
    window.show()
    mark("create SummaryApp end")
    mark("mainloop begin")
    return app.exec()
