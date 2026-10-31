"""Reusable native components. No application, file IO or timers at import."""
from math import ceil
from PySide6.QtCore import Qt, QSize, QTimer
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (QAbstractItemView, QComboBox, QFrame, QHBoxLayout,
    QHeaderView, QLabel, QPushButton, QTableWidget, QTableWidgetItem, QToolButton,
    QVBoxLayout, QWidget)
from .demo import TableData


def text(value, role=None, *, wrap=False):
    control = QLabel(value)
    if role:
        control.setObjectName(role)
    control.setWordWrap(wrap)
    return control


def push(caption, role=None, *, icon=None):
    control = QPushButton(caption)
    if role:
        control.setObjectName(role)
    if icon is not None:
        control.setIcon(icon)
        control.setIconSize(QSize(22, 22))
    control.setCursor(Qt.CursorShape.PointingHandCursor)
    control.setMinimumHeight(32)
    control.setAccessibleName(caption)
    return control


class Disclosure(QWidget):
    def __init__(self, title, icons, parent=None):
        super().__init__(parent)
        self.icons = icons
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        self.toggle = QToolButton()
        self.toggle.setObjectName("Disclosure")
        self.toggle.setText(title)
        self.toggle.setCheckable(True)
        self.toggle.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.toggle.setIcon(icons.get("next", size=18))
        self.toggle.setIconSize(QSize(18, 18))
        self.toggle.setMinimumHeight(36)
        self.toggle.setAccessibleName(title)
        self.toggle.setCursor(Qt.CursorShape.PointingHandCursor)
        self.body = QWidget()
        self.body.setVisible(False)
        layout.addWidget(self.toggle, alignment=Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(self.body)
        self.toggle.toggled.connect(self.set_open)

    def set_open(self, opened):
        self.body.setVisible(opened)
        self.toggle.setIcon(self.icons.get("down" if opened else "next", size=18))


class TablePanel(QWidget):
    PAGE_SIZE = 100

    def __init__(self, icons, parent=None):
        super().__init__(parent)
        self.page = 0
        self.data = TableData((), ())
        self.minimum_column_widths = {}
        self.fit_timer = QTimer(self)
        self.fit_timer.setSingleShot(True)
        self.fit_timer.timeout.connect(self.fit_columns)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.table.verticalHeader().hide()
        self.table.verticalHeader().setDefaultSectionSize(38)
        header = self.table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setMinimumSectionSize(100)
        header.setFixedHeight(40)
        layout.addWidget(self.table, 1)
        footer = QHBoxLayout()
        footer.setContentsMargins(0, 0, 0, 0)
        self.page_size = QComboBox()
        self.page_size.addItem("每页 100 行")
        self.page_size.setFixedWidth(130)
        self.page_size.setToolTip("演示与现有工作台均按每页100条数据分页；合计为展示栏。")
        footer.addWidget(self.page_size)
        footer.addStretch()
        self.count = text("第 1 / 1 页", "Small")
        footer.addWidget(self.count)
        self.previous = push("", icon=icons.get("previous", size=18))
        self.next = push("", icon=icons.get("next", size=18))
        for button, caption in ((self.previous, "上一页"), (self.next, "下一页")):
            button.setFixedSize(36, 32)
            button.setAccessibleName(caption)
            button.setToolTip(caption)
        self.previous.clicked.connect(lambda: self.move(-1))
        self.next.clicked.connect(lambda: self.move(1))
        footer.addWidget(self.previous)
        footer.addWidget(self.next)
        layout.addLayout(footer)

    def set_data(self, data):
        self.data, self.page = data, 0
        self.table.setColumnCount(len(data.headers))
        self.table.setHorizontalHeaderLabels(data.headers)
        self.render()
        self.fit_columns()
        self.fit_timer.start(0)

    def move(self, delta):
        self.page = max(0, min(max(0, ceil(len(self.data.rows) / self.PAGE_SIZE) - 1), self.page + delta))
        self.render()

    def render(self):
        pages = max(1, ceil(len(self.data.rows) / self.PAGE_SIZE))
        rows = self.data.rows[self.page * self.PAGE_SIZE:(self.page + 1) * self.PAGE_SIZE]
        show_footer = self.data.footer is not None and self.page == pages - 1
        self.table.setRowCount(len(rows) + int(show_footer))
        for index, row in enumerate((*rows, self.data.footer) if show_footer else rows):
            for column, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setToolTip(str(value))
                if show_footer and index == len(rows):
                    font = QFont(self.table.font())
                    font.setBold(True)
                    item.setFont(font)
                    item.setBackground(QColor("#F0F2F6"))
                    item.setData(Qt.ItemDataRole.UserRole, "presentation-footer")
                self.table.setItem(index, column, item)
        self.count.setText(f"第 {self.page + 1} / {pages} 页")
        self.previous.setEnabled(self.page > 0)
        self.next.setEnabled(self.page < pages - 1)

    def fit_columns(self):
        number = self.table.columnCount()
        if not number:
            return
        available = self.table.viewport().width()
        weights = (1.05, 1.0, 0.75) if number == 3 else (1,) * number
        metrics = self.table.horizontalHeader().fontMetrics()
        minimums = [max(110, metrics.horizontalAdvance(header) + 30) for header in self.data.headers]
        for index, width in self.minimum_column_widths.items():
            if 0 <= index < number:
                minimums[index] = max(minimums[index], width)
        for row in self.data.rows[:10]:
            for index, value in enumerate(row):
                minimums[index] = max(minimums[index], min(280, metrics.horizontalAdvance(str(value)) + 24))
        widths = [max(minimums[index], int(available * weight / sum(weights)))
                  for index, weight in enumerate(weights[:-1])]
        widths.append(max(minimums[-1], available - sum(widths)))
        for index, width in enumerate(widths):
            self.table.setColumnWidth(index, width)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "fit_timer"):
            self.fit_timer.start(0)  # Fit after child viewport/layout updates.
