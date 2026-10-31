"""Source SVG icons rendered by Qt at the actual device density, fully local."""
from pathlib import Path
from PySide6.QtCore import QByteArray, Qt
from PySide6.QtGui import QIcon, QPixmap, QPainter
from PySide6.QtSvg import QSvgRenderer

PROJECT = Path(__file__).resolve().parents[2]
NAMES = {"workbench": "home", "history": "clock", "rules": "file-text",
         "help": "book-open", "feedback": "message-square", "about": "info",
         "settings": "settings", "offline": "settings", "calendar": "calendar", "previous": "chevron-left",
         "next": "chevron-right", "down": "chevron-down", "up": "chevron-up",
         "external": "arrow-up-right", "file": "file-text", "remove": "x"}


class Icons:
    def __init__(self):
        self.cache = {}

    def get(self, name, *, active=False, size=22):
        key = (name, active, size)
        if key not in self.cache:
            color = "#0071E3" if active else "#667085"
            svg = (PROJECT / "assets/ui" / (NAMES[name] + ".svg")).read_bytes().replace(b"currentColor", color.encode())
            renderer = QSvgRenderer(QByteArray(svg))
            if not renderer.isValid():
                raise ValueError(f"Invalid bundled icon: {name}")
            icon = QIcon()
            for ratio in (1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4):
                pixels = round(size * ratio)
                pixmap = QPixmap(pixels, pixels)
                pixmap.fill(Qt.GlobalColor.transparent)
                painter = QPainter(pixmap)
                renderer.render(painter)
                painter.end()
                pixmap.setDevicePixelRatio(ratio)
                icon.addPixmap(pixmap, QIcon.Mode.Normal)
            self.cache[key] = icon
        return self.cache[key]

    def brand(self):
        return QIcon(str(PROJECT / "assets/app-icon.png"))
