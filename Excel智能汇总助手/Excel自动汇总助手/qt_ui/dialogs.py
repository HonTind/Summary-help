"""Native dialogs; persistence remains behind the existing business interfaces."""
import json
from pathlib import Path
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDialog, QDialogButtonBox,
    QFormLayout, QHBoxLayout, QLineEdit, QListWidget, QListWidgetItem, QMessageBox,
    QScrollArea, QSpinBox, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)
from ..analysis import SheetRules
from .. import field_mapper
from .widgets import push, text


def selection(value):
    value = value.strip()
    if value.startswith("#"):
        if not value[1:].isdigit():
            raise ValueError("列序号请使用 #1、#2 等格式。")
        return int(value[1:])
    return value


def display_field(value):
    return f"#{value}" if type(value) is int else str(value)


def fit_dialog(dialog, width=620, height=500):
    available = dialog.screen().availableGeometry().size()
    dialog.resize(min(width, available.width() - 40), min(height, available.height() - 60))


class RulesDialog(QDialog):
    def __init__(self, parent, rules, title="字段与工作表规则"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.value = None
        layout = QVBoxLayout(self)
        layout.addWidget(text("键按列表顺序组合，最多8项。字段填写表头名称或 #列号。", wrap=True))
        self.keys = QListWidget()
        for key in rules.field_selections():
            item = QListWidgetItem(display_field(key))
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
            self.keys.addItem(item)
        layout.addWidget(self.keys, 1)
        actions = QHBoxLayout()
        for caption, callback in (("添加", self.add_key), ("移除", self.remove_key),
                                   ("上移", lambda: self.move_key(-1)), ("下移", lambda: self.move_key(1))):
            button = push(caption)
            button.clicked.connect(callback)
            actions.addWidget(button)
        layout.addLayout(actions)
        form = QFormLayout()
        self.amount = QLineEdit(display_field(rules.amount_column))
        self.sheet = QLineEdit(rules.sheet or "")
        self.sheet.setPlaceholderText("留空使用默认工作表")
        self.header = QSpinBox()
        self.header.setRange(1, 1048576)
        self.header.setValue(rules.header_row)
        self.date = QLineEdit(display_field(rules.date_column) if rules.date_column is not None else "")
        for caption, control in (("金额字段", self.amount), ("工作表", self.sheet),
                                 ("表头行", self.header), ("日期字段（留空关闭）", self.date)):
            form.addRow(caption, control)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.validate)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        fit_dialog(self, height=540)

    def add_key(self):
        if self.keys.count() >= 8:
            return
        item = QListWidgetItem("新字段")
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
        self.keys.addItem(item)
        self.keys.setCurrentItem(item)
        self.keys.editItem(item)

    def remove_key(self):
        if self.keys.currentRow() >= 0:
            self.keys.takeItem(self.keys.currentRow())

    def move_key(self, delta):
        row = self.keys.currentRow()
        if 0 <= row + delta < self.keys.count():
            item = self.keys.takeItem(row)
            self.keys.insertItem(row + delta, item)
            self.keys.setCurrentRow(row + delta)

    def validate(self):
        try:
            keys = tuple(selection(self.keys.item(i).text()) for i in range(self.keys.count()))
            if not keys:
                raise ValueError("至少需要一个键字段。")
            self.value = SheetRules(key_column=keys[0], amount_column=selection(self.amount.text()),
                sheet=self.sheet.text().strip() or None, header_row=self.header.value(),
                key_fields=keys if len(keys) > 1 else None,
                date_column=selection(self.date.text()) if self.date.text().strip() else None)
        except ValueError as exc:
            QMessageBox.warning(self, "规则无效", str(exc))
            return
        self.accept()


class FileSelectionDialog(QDialog):
    def __init__(self, parent, paths):
        super().__init__(parent)
        self.setWindowTitle("选择参与汇总的文件")
        self.paths, self.value = paths, None
        layout = QVBoxLayout(self)
        self.files = QListWidget()
        for path in paths:
            item = QListWidgetItem(Path(path).name)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            self.files.addItem(item)
        layout.addWidget(self.files, 1)
        row = QHBoxLayout()
        for caption, mode in (("全选", "all"), ("全不选", "none"), ("反选", "invert")):
            button = push(caption)
            button.clicked.connect(lambda _checked=False, m=mode: self.select(m))
            row.addWidget(button)
        layout.addLayout(row)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.confirm)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        fit_dialog(self)

    def select(self, mode):
        for index in range(self.files.count()):
            item = self.files.item(index)
            checked = mode == "all" or (mode == "invert" and item.checkState() != Qt.CheckState.Checked)
            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)

    def confirm(self):
        self.value = [path for i, path in enumerate(self.paths) if self.files.item(i).checkState() == Qt.CheckState.Checked]
        if not self.value:
            QMessageBox.warning(self, "选择文件", "至少选择一个文件。")
            return
        self.accept()


class MappingDialog(QDialog):
    def __init__(self, parent, path, current):
        super().__init__(parent)
        self.setWindowTitle("字段映射 · " + Path(path).name)
        self.value = None
        layout = QVBoxLayout(self)
        ok, headers = field_mapper.get_file_headers(path)
        if not ok:
            raise ValueError(headers)
        names = [row["name"] for row in field_mapper.get_field_config()]
        self.names, self.controls = names, []
        layout.addWidget(text("选择各列含义，前两个字段必填。默认仅对本次任务生效。", wrap=True))
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        body = QWidget()
        form = QFormLayout(body)
        effective = dict(field_mapper.get_all_mappings(), **current)
        for index, header in headers:
            combo = QComboBox()
            combo.addItems(["忽略", *names])
            guess = effective.get(header)
            if guess in names:
                combo.setCurrentText(guess)
            elif header:
                for config in field_mapper.get_field_config():
                    if header == config["name"] or any(word in header for word in config["keywords"]):
                        combo.setCurrentText(config["name"])
                        break
            form.addRow(f"#{index + 1} {header or '空表头'}", combo)
            self.controls.append((header, combo))
        scroll.setWidget(body)
        layout.addWidget(scroll, 1)
        self.remember = QCheckBox("记住映射（下次仍生效）")
        layout.addWidget(self.remember)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.validate)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        fit_dialog(self)

    def validate(self):
        chosen = [(header, control.currentText()) for header, control in self.controls if control.currentIndex() > 0]
        if any(not header for header, _ in chosen) or len({h for h, _ in chosen}) != len(chosen):
            QMessageBox.warning(self, "字段映射", "不能映射空表头或同名的多个表头，请先在 Excel 修正。")
            return
        if len({name for _, name in chosen}) != len(chosen) or any(name not in dict((n, h) for h, n in chosen) for name in self.names[:2]):
            QMessageBox.warning(self, "字段映射", "前两个字段必填，同一字段只能选择一列。")
            return
        self.value = dict(chosen)
        self.accept()


class FieldConfigDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("常规汇总字段规则")
        self.value = None
        layout = QVBoxLayout(self)
        layout.addWidget(text("第1个字段用于分组，第2个字段用于求和。关键词以逗号分隔。", wrap=True))
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["字段名称（顺序有意义）", "识别关键词"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 200)
        for row in field_mapper.get_field_config():
            self.add_row(row["name"], ", ".join(row["keywords"]))
        layout.addWidget(self.table, 1)
        buttons_row = QHBoxLayout()
        add = push("添加字段")
        add.clicked.connect(lambda: self.add_row())
        remove = push("移除字段")
        remove.clicked.connect(lambda: self.table.removeRow(self.table.currentRow()) if self.table.currentRow() >= 2 else None)
        buttons_row.addWidget(add)
        buttons_row.addWidget(remove)
        layout.addLayout(buttons_row)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.validate)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        fit_dialog(self)

    def add_row(self, name="", keywords=""):
        row = self.table.rowCount()
        self.table.insertRow(row)
        self.table.setItem(row, 0, QTableWidgetItem(name))
        self.table.setItem(row, 1, QTableWidgetItem(keywords))

    def validate(self):
        result = []
        for row in range(self.table.rowCount()):
            name = self.table.item(row, 0).text().strip()
            words = [w.strip() for w in self.table.item(row, 1).text().replace("，", ",").split(",") if w.strip()]
            if not name or not words or name in {entry["name"] for entry in result}:
                QMessageBox.warning(self, "字段规则", "字段名称与关键词不能为空，字段名称不能重复。")
                return
            result.append({"name": name, "keywords": words})
        if len(result) < 2:
            return
        self.value = result
        self.accept()


def validate_existing_config(path):
    """Do not silently overwrite a damaged legacy configuration."""
    path = Path(path)
    if path.exists():
        with path.open(encoding="utf-8") as handle:
            value = json.load(handle)
        if not isinstance(value, (dict, list)):
            raise ValueError("字段配置结构无效，原文件已保留。")
