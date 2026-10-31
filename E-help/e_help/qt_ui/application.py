"""Production Qt views. The main thread alone owns widget and history updates."""
from copy import deepcopy
from dataclasses import replace
import datetime
import os
from pathlib import Path
import queue

from PySide6.QtCore import QTimer, QUrl, Qt
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (QApplication, QComboBox, QDialog, QFileDialog, QHBoxLayout,
    QInputDialog, QLineEdit, QMessageBox, QSizePolicy, QVBoxLayout, QWidget)
from ..analysis import SheetRules
from ..config import APP_FULL_NAME, OUTPUT_FOLDER_NAME
from ..edition_policy import edition_status
from .. import field_mapper, history
from ..excel_reader import _get_excel_files
from ..job_runner import JobRunner
from ..template_store import make_template
from .backend import Backend, Request
from .demo import TableData
from .dialogs import (FieldConfigDialog, FileSelectionDialog, MappingDialog, RulesDialog,
                      display_field, selection, validate_existing_config)
from .widgets import TablePanel, push, text
from .window import PreviewWindow


class ApplicationWindow(PreviewWindow):
    def __init__(self, backend, *, initialize=True, available_size=None):
        self.production_backend = backend
        self.right_files = []
        self.left_rules, self.right_rules = SheetRules(), None
        self.mappings, self.output_path, self.input_folder = {}, "", ""
        self._templates, self._records = {}, []
        self._configured = False
        self.runner = JobRunner()
        self._terminal, self._job_request, self._job_result = None, None, None
        self._phase, self._close_pending = "preview", False
        super().__init__(available_size=available_size, populate_demo=False)
        self.backend = backend
        self.setWindowTitle(APP_FULL_NAME)
        self.selected_files = []
        self.choose_files_button.setToolTip("选择本地 Excel 文件；输入始终只读。")
        self.large_sample.hide()
        self.save_template.setText("保存月度模板")
        self.advanced_form.itemAt(0).widget().setText("工作表、表头行和组合键可单独配置。")
        self.extra_buttons = []
        for caption, callback in (("字段与工作表规则", lambda: self.edit_rules(False)),
                                   ("右侧规则（默认同左）", lambda: self.edit_rules(True)),
                                   ("常规字段映射", self.edit_mapping),
                                   ("管理常规字段规则", self.edit_field_config)):
            button = push(caption, "Link")
            button.clicked.connect(callback)
            self.advanced_form.addWidget(button)
            self.controls.append(button)
            self.extra_buttons.append(button)
        self.output_prefix = QLineEdit("月度汇总")
        self.output_prefix.setPlaceholderText("导出文件名前缀")
        self.output_prefix.textChanged.connect(self.invalidate)
        self.prefix_label = text("输出前缀", "Small")
        self.advanced_form.addWidget(self.prefix_label)
        self.advanced_form.addWidget(self.output_prefix)
        self.controls.append(self.output_prefix)
        self.tolerance = QLineEdit("0")
        self.tolerance.setPlaceholderText("金额容差（默认0）")
        self.tolerance.textChanged.connect(self.invalidate)
        self.advanced_form.addWidget(self.tolerance)
        self.controls.append(self.tolerance)
        self.choose_right = push("选择右侧账单", "Link")
        self.choose_right.clicked.connect(self.choose_right_files)
        input_layout = self.file_card.parentWidget().layout()
        left_index = next(index for index in range(input_layout.count())
                          if input_layout.itemAt(index).layout() is not None
                          and input_layout.itemAt(index).layout().indexOf(self.choose_files_button) >= 0)
        actions = input_layout.itemAt(left_index).layout()
        # Remove the trailing spacer so both input buttons fill the same column.
        actions.takeAt(actions.count() - 1)
        input_layout.insertWidget(left_index + 1, self.choose_right)
        self.controls.append(self.choose_right)
        self.output_mode = QComboBox()
        self.output_mode.addItems(["默认输出目录", "输入文件夹", "自定义目录"])
        self.custom_output = ""
        self.advanced_form.addWidget(text("输出位置", "Small"))
        self.advanced_form.addWidget(self.output_mode)
        self.controls.append(self.output_mode)
        self.custom_output_button = push("选择自定义目录", "Link")
        self.custom_output_button.clicked.connect(self.choose_output_directory)
        self.advanced_form.addWidget(self.custom_output_button)
        self.controls.append(self.custom_output_button)
        self.output_mode.currentIndexChanged.connect(self.save_output_preference)
        self.cancel_button = push("取消")
        self.cancel_button.clicked.connect(self.cancel_job)
        self.footer.layout().insertWidget(2, self.cancel_button)
        self.cancel_button.hide()
        self.poll_timer = QTimer(self)
        self.poll_timer.setInterval(40)
        self.poll_timer.timeout.connect(self.poll_job)
        self.month.setCurrentText(datetime.date.today().strftime("%Y-%m"))
        self.template.setCurrentText("标准月度规则")
        self._configured = True
        startup_warning = ""
        if initialize:
            try:
                backend.initialize()
                self.refresh_templates()
                preferences = backend.preferences()
                self.output_mode.blockSignals(True)
                self.output_mode.setCurrentIndex({"default": 0, "input": 1, "custom": 2}.get(preferences.get("output_mode"), 0))
                self.custom_output = preferences.get("output_dir", "") if isinstance(preferences.get("output_dir", ""), str) else ""
                self.output_mode.blockSignals(False)
            except (OSError, ValueError) as exc:
                startup_warning = f"配置未加载：{exc}；原文件已保留。"
        self.select_mode("regular")
        self.info.setText("选择输入文件后预览，再导出结果。全部操作在本机完成。")
        self.status.setText(startup_warning or "就绪 · 输入只读，导出使用新文件名。")
        self.export_button.setToolTip("导出当前预览的完整结果，不覆盖已有文件。")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "footer_hint"):
            self.footer_hint.setText("本地处理 · 导出不覆盖" if self.width() < 1000 else "先预览再导出；输入文件保持原样。")

    def invalidate(self, *_):
        if hasattr(self, "export_button") and not self.busy:
            self.result, self.output_path = None, ""
            self.export_button.setEnabled(False)
            self.info.setText("输入或规则已更改，请重新预览。")
            for panel in self.table_panels:
                panel.set_data(TableData(panel.data.headers, ()))

    def date_toggled(self, checked):
        self.date_field.setVisible(self.mode == "monthly" and checked)
        self.date_hint.setText("开启时，只处理所选月份；排除行可在来源查看。" if checked else "关闭时，月份仅作标签。")
        self.invalidate()

    def refresh_files(self):
        while self.file_layout.count():
            item = self.file_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        left = self.selected_files
        groups = [("左侧", left), ("右侧", self.right_files)] if self.mode == "compare" else [("", left)]
        for side, files in groups:
            if side:
                self.file_layout.addWidget(text(f"{side} · {len(files)}个文件", "Small"))
            for path in files[:4]:
                row = QWidget()
                layout = QHBoxLayout(row)
                layout.setContentsMargins(0, 4, 0, 4)
                layout.setSpacing(10)
                layout.addWidget(self.image_label("file"))
                name = text(Path(path).name, wrap=True)
                name.setToolTip(str(path))
                layout.addWidget(name, 1)
                self.file_layout.addWidget(row)
            if len(files) > 4:
                self.file_layout.addWidget(text(f"另有 {len(files) - 4} 个文件", "Small"))
            if not files:
                self.file_layout.addWidget(text("尚未选择文件", "Small"))
        self.file_count.setText(f"{len(left) + (len(self.right_files) if self.mode == 'compare' else 0)}个文件")

    def choose_samples(self):
        if self.busy:
            return
        if self.mode == "regular":
            folder = QFileDialog.getExistingDirectory(self, "选择输入文件夹", self.input_folder)
            if not folder:
                return
            paths = [str(Path(folder) / name) for name in _get_excel_files(folder)]
            if not paths:
                QMessageBox.information(self, "选择文件", "文件夹中没有可用 Excel 文件。")
                return
            dialog = FileSelectionDialog(self, paths)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                return
            self.input_folder, self.selected_files = folder, dialog.value
            self.mappings = {}
        else:
            paths, _ = QFileDialog.getOpenFileNames(self, "选择左侧账单" if self.mode == "compare" else "选择本月账单",
                "", "Excel (*.xlsx *.xlsm *.xls)")
            if not paths:
                return
            self.selected_files = paths
        self.refresh_files()
        self.invalidate()

    def choose_right_files(self):
        if self.busy:
            return
        paths, _ = QFileDialog.getOpenFileNames(self, "选择右侧账单", "", "Excel (*.xlsx *.xlsm *.xls)")
        if paths:
            self.right_files = paths
            self.refresh_files()
            self.invalidate()

    def rules(self):
        first = selection(self.key.currentText())
        keys = self.left_rules.key_fields
        if keys:
            keys = (first, *keys[1:])
        return replace(self.left_rules, key_column=first, key_fields=keys,
            amount_column=selection(self.amount.currentText()),
            date_column=selection(self.date_field.text()) if self.mode == "monthly" and self.date_check.isChecked() else None)

    def request(self):
        if not getattr(self, "_configured", False):
            return None
        rules = self.rules()
        return Request(self.mode, tuple(self.selected_files), tuple(self.right_files), rules,
            self.right_rules or rules, self.month.currentText().strip(), self.template.currentText().strip(),
            self.output_prefix.text().strip(), "customer" if self.compare_selector.currentIndex() == 0 else "document",
            self.tolerance.text().strip(), deepcopy(self.mappings))

    def apply_rules(self, rules):
        self.left_rules = rules
        self.key.setCurrentText(display_field(rules.field_selections()[0]))
        self.amount.setCurrentText(display_field(rules.amount_column))
        self.date_check.setChecked(rules.date_column is not None)
        self.date_field.setText(display_field(rules.date_column) if rules.date_column is not None else "日期")
        self.invalidate()

    def edit_rules(self, right=False):
        if self.busy:
            return
        try:
            dialog = RulesDialog(self, self.right_rules or self.rules() if right else self.rules(),
                                 "右侧账单规则" if right else "月度/左侧账单规则")
            if dialog.exec() == QDialog.DialogCode.Accepted:
                if right:
                    self.right_rules = replace(dialog.value, date_column=None)
                    self.invalidate()
                else:
                    self.apply_rules(dialog.value)
        except (OSError, ValueError) as exc:
            QMessageBox.warning(self, "规则", str(exc))

    def refresh_templates(self):
        self._templates = self.production_backend.templates()
        selected = self.template.currentText()
        self.template.blockSignals(True)
        self.template.clear()
        self.template.addItems(["标准月度规则", *(name for name in self._templates if name != "标准月度规则")])
        self.template.setCurrentText(selected)
        self.template.blockSignals(False)

    def load_memory_template(self, *_):
        if self.busy or not self._configured:
            return
        try:
            self.refresh_templates()
            template = self._templates.get(self.template.currentText())
            if template:
                self.apply_rules(SheetRules(**template["rules"]))
                self.output_prefix.setText(template["output_prefix"])
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "模板", f"无法读取模板，原文件保留：{exc}")

    def save_memory_template(self):
        if self.busy:
            return
        try:
            template = make_template(self.template.currentText().strip(), self.rules(), self.output_prefix.text().strip())
            templates = self.production_backend.templates()
            allow = template["name"] in templates
            if allow and QMessageBox.question(self, "更新同名模板", "更新这份月度模板？旧账单名单不会保存。") != QMessageBox.StandardButton.Yes:
                return
            self.production_backend.save_template(template, allow_update=allow)
            self.refresh_templates()
            self.info.setText("月度模板已保存在本机，未保存账单名单或内容。")
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "模板未保存", str(exc))

    def choose_output_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录", self.custom_output)
        if directory:
            self.custom_output = directory
            self.output_mode.setCurrentIndex(2)
            self.save_output_preference()

    def save_output_preference(self, *_):
        if not self._configured or self.busy:
            return
        try:
            self.production_backend.save_preferences(("default", "input", "custom")[self.output_mode.currentIndex()], self.custom_output)
        except (ValueError, OSError) as exc:
            self.status.setText(f"输出设置未保存，原文件保留：{exc}")

    def output_directory(self):
        mode = self.output_mode.currentIndex()
        if mode == 0:
            path = self.production_backend.base_dir / OUTPUT_FOLDER_NAME
            path.mkdir(parents=True, exist_ok=True)
            return path
        if mode == 1:
            parents = {Path(p).parent for p in self.selected_files}
            if len(parents) != 1:
                raise ValueError("输入文件位于多个文件夹，请选择自定义输出目录。")
            return next(iter(parents))
        if not self.custom_output or not Path(self.custom_output).is_dir():
            raise ValueError("请先选择有效的自定义输出目录。")
        return Path(self.custom_output)

    def preview(self):
        if self.busy:
            return
        try:
            request = self.request()
            if not request.paths:
                raise ValueError("请先选择左侧账单。" if request.mode == "compare"
                                 else "请先选择包含 Excel 文件的输入文件夹。" if request.mode == "regular"
                                 else "请先选择本月账单文件。")
            if request.mode == "compare" and not request.right_paths:
                raise ValueError("请先选择右侧账单。")
        except ValueError as exc:
            QMessageBox.warning(self, "无法预览", str(exc))
            return
        self.invalidate()
        self.start_job(request, "preview", lambda cancel, progress: self.production_backend.preview(request, cancel, progress))

    def explain_export(self):
        if self.busy or self.result is None:
            return
        result = self.result
        if not result.complete and QMessageBox.question(self, "结果不完整", "结果包含异常或重复歧义，仍导出供核查？\n报告会保留异常，不代表全部一致。") != QMessageBox.StandardButton.Yes:
            return
        try:
            directory = self.output_directory()
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "输出位置", str(exc))
            return
        self.start_job(result.request, "export", lambda cancel, progress: self.production_backend.export(result, directory, cancel), result=result)

    def start_job(self, request, phase, operation, *, result=None):
        if self.busy or self.runner.busy:
            return False
        self._job_request, self._phase, self._job_result = deepcopy(request), phase, result
        self._terminal = None
        self.busy = True
        if not self.runner.start(operation, accept_completed=phase == "export"):
            self.busy = False
            return False
        self.set_busy_controls(True)
        self.info.setText("正在预览…" if phase == "preview" else "正在导出完整结果…")
        self.poll_timer.start()
        return True

    def set_busy_controls(self, busy):
        for control in (*self.controls, *self.mode_buttons.values(), self.preview_button):
            control.setEnabled(not busy)
        for page in ("workbench", "history", "rules"):
            self.nav_buttons[page].setEnabled(not busy)
        self.export_button.setEnabled(not busy and self.result is not None)
        self.cancel_button.setVisible(busy)
        self.cancel_button.setEnabled(busy and self._job_request.mode != "regular")
        self.cancel_button.setToolTip("常规汇总需等待结束" if self._job_request.mode == "regular" else "取消当前任务")

    def cancel_job(self):
        if self.busy and self._job_request.mode != "regular":
            self.runner.cancel()
            self.cancel_button.setEnabled(False)
            self.info.setText("正在取消，等待文件资源释放…")

    def poll_job(self):
        while True:
            try:
                kind, value = self.runner.events.get_nowait()
            except queue.Empty:
                break
            if kind == "progress":
                self.status.setText(str(value))
            else:
                self._terminal = (kind, value)
        if self.runner.busy:
            return
        # The worker can enqueue a terminal event between the first drain and exit.
        while True:
            try:
                kind, value = self.runner.events.get_nowait()
            except queue.Empty:
                break
            if kind != "progress":
                self._terminal = (kind, value)
        if self._terminal is None:
            return
        terminal, self._terminal = self._terminal, None
        if terminal[0] == "done" and self._phase == "preview" and self.runner.cancel_event.is_set():
            terminal = ("cancelled", "任务已取消")
        self.poll_timer.stop()
        self.finish_job(*terminal)

    def finish_job(self, kind, value):
        request, phase = self._job_request, self._phase
        status, result, output = "failed", self._job_result, ""
        if kind == "done":
            if phase == "preview":
                result = value
                self.show_result(result)
                status = "previewed" if result.complete else "incomplete"
            else:
                output = str(value)
                self.output_path = output
                self.info.setText("已导出完整结果：" + Path(output).name)
                self.status.setText("保存完成。可在任务记录中打开结果文件或所在目录。")
                status = "success"
        else:
            status = "cancelled" if kind == "cancelled" else "failed"
            if phase == "preview":
                self.result = None
            self.info.setText("任务已取消。" if kind == "cancelled" else f"任务失败：{value}")
            self.status.setText("文件资源已释放，可以调整后重试。")
        try:
            recorded = self.production_backend.record(request, status, phase, result=result, output=output)
        except (OSError, ValueError):
            recorded = False
        if not recorded:
            self.status.setText(self.status.text() + " 任务记录未保存；已导出的文件不受影响。")
        self.busy = False
        self.set_busy_controls(False)
        if self._close_pending:
            self.close()

    def show_result(self, result):
        self.result = result
        for panel, data in zip(self.table_panels, (result.preview.summary, result.preview.sources, result.preview.issues)):
            panel.set_data(data)
        self.info.setText(result.preview.message)
        self.status.setText(f"预览完成 · {len(result.preview.summary.rows)}条结果 · {len(result.preview.issues.rows)}项异常。")

    def select_mode(self, mode):
        if self.busy or mode not in self.MODES:
            return
        super().select_mode(mode)
        if not self._configured:
            return
        monthly, compare, regular = mode == "monthly", mode == "compare", mode == "regular"
        for control in (self.template, self.template.field_label, self.save_template, self.output_prefix, self.prefix_label):
            control.setVisible(monthly)
        for control in (self.key, self.key.field_label, self.amount, self.amount.field_label):
            control.setVisible(not regular)
        self.choose_right.setVisible(compare)
        self.tolerance.setVisible(compare)
        self.extra_buttons[0].setVisible(not regular)
        self.extra_buttons[1].setVisible(compare)
        for control in self.extra_buttons[2:]:
            control.setVisible(regular)
        self.choose_files_button.setText("选择文件夹" if regular else "选择左侧账单" if compare else "更换文件")
        for spacer, height in self.form_spacers:
            spacer.changeSize(0, height if monthly else 6, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
        self.settings_form.invalidate()
        self.refresh_files()

    def edit_mapping(self):
        if self.busy or not self.selected_files:
            return
        names = [Path(path).name for path in self.selected_files]
        chosen, ok = QInputDialog.getItem(self, "字段映射", "选择文件", names, 0, False)
        if not ok:
            return
        path = self.selected_files[names.index(chosen)]
        try:
            dialog = MappingDialog(self, path, self.mappings.get(chosen, {}))
            if dialog.exec() == QDialog.DialogCode.Accepted:
                if dialog.remember.isChecked():
                    self.production_backend.preferences()  # Validate before legacy writer.
                    for header, name in dialog.value.items():
                        field_mapper.save_field_mapping(str(self.production_backend.config_dir), header, name)
                self.mappings[chosen] = dialog.value
                self.invalidate()
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "字段映射", str(exc))

    def edit_field_config(self):
        if self.busy:
            return
        try:
            path = self.production_backend.config_dir / field_mapper.FIELD_CONFIG_FILENAME
            validate_existing_config(path)
            dialog = FieldConfigDialog(self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self.production_backend.config_dir.mkdir(parents=True, exist_ok=True)
                field_mapper.save_field_config(str(self.production_backend.config_dir), dialog.value)
                self.mappings = {}
                self.invalidate()
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "字段规则", str(exc))

    def build_history(self):
        card = self.card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 20, 20, 20)
        self.history_warning = text("", "Small", wrap=True)
        self.history_kind = QComboBox()
        self.history_kind.addItems(["全部任务", *history.TASK_TYPES.values()])
        self.history_kind.currentIndexChanged.connect(self.refresh_history)
        layout.addWidget(self.history_kind)
        layout.addWidget(self.history_warning)
        self.history_panel = TablePanel(self.icons)
        layout.addWidget(self.history_panel, 1)
        row = QHBoxLayout()
        for caption, callback in (("刷新", self.refresh_history), ("打开结果", lambda: self.open_history(False)),
                                   ("打开文件位置", lambda: self.open_history(True)),
                                   ("删除选中记录", self.delete_history), ("清空记录", self.clear_history)):
            button = push(caption)
            button.clicked.connect(callback)
            row.addWidget(button)
        layout.addLayout(row)
        layout.addWidget(text("删除只影响任务记录，Excel 文件始终保留。", "Small"))
        return card

    def refresh_history(self, *_):
        if not self._configured or self.busy:
            return
        records, warning = history.read_history()
        kind = None if self.history_kind.currentIndex() == 0 else list(history.TASK_TYPES)[self.history_kind.currentIndex() - 1]
        self._records = list(reversed(history.filter_records(records, kind)))
        self.history_warning.setText(warning)
        self.history_panel.minimum_column_widths = {0: 210, 4: 300}
        self.history_panel.set_data(TableData(("时间", "类型", "状态", "文件数", "结果文件"),
            tuple((r.get("time", ""), history.type_label(r), history.status_label(r), r.get("file_count", ""), r.get("output_file", "")) for r in self._records)))

    def selected_record(self):
        row = self.history_panel.table.currentRow()
        index = self.history_panel.page * self.history_panel.PAGE_SIZE + row
        return self._records[index] if row >= 0 and index < len(self._records) else None

    def open_history(self, folder):
        record = self.selected_record()
        path = history.output_path(record) if record else ""
        if not path or not Path(path).is_file():
            QMessageBox.information(self, "打开结果", "请选择仍存在的已导出结果。")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(path).parent) if folder else path))

    def delete_history(self):
        record = self.selected_record()
        if record and QMessageBox.question(self, "删除记录", "仅删除选中记录？Excel 文件保留。") == QMessageBox.StandardButton.Yes:
            if not history.delete_record(record["id"]):
                QMessageBox.warning(self, "记录未删除", "记录文件无法更新，原文件保留。")
            self.refresh_history()

    def clear_history(self):
        if QMessageBox.question(self, "清空记录", "仅清空任务记录？所有 Excel 文件保留。") == QMessageBox.StandardButton.Yes:
            if not history.clear_records_safe():
                QMessageBox.warning(self, "记录未清空", "记录文件无法更新，原文件保留。")
            self.refresh_history()

    def build_rules(self):
        card = self.simple_page("规则与模板", ["常规规则用于自动识别；月度模板保存显式字段规则。对账左右两侧可分别设置。"])
        for index, (caption, callback) in enumerate((("管理常规字段规则", self.edit_field_config),
            ("月度模板", lambda: self.select_mode("monthly")), ("管理已记住的字段映射", self.manage_mappings),
            ("重置常规字段规则", self.reset_field_config))):
            button = push(caption, "Link")
            button.clicked.connect(callback)
            card.layout().insertWidget(index + 2, button, alignment=Qt.AlignmentFlag.AlignLeft)
        return card

    def reset_field_config(self):
        if self.busy or QMessageBox.question(self, "重置规则", "恢复原常规字段默认规则？用户自定义规则文件会移除，Excel 文件保留。") != QMessageBox.StandardButton.Yes:
            return
        try:
            validate_existing_config(self.production_backend.config_dir / field_mapper.FIELD_CONFIG_FILENAME)
            field_mapper.reset_field_config(str(self.production_backend.config_dir))
            self.mappings = {}
            self.invalidate()
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "规则未重置", str(exc))

    def manage_mappings(self):
        mappings = field_mapper.get_all_mappings()
        entries = [f"{header} → {name}" for header, name in mappings.items()]
        if not entries:
            QMessageBox.information(self, "字段映射", "没有已记住的字段映射。")
            return
        selected, ok = QInputDialog.getItem(self, "管理已记住的映射", "选择需要移除的映射", entries, 0, False)
        if ok and QMessageBox.question(self, "移除映射", "移除这条已记住的映射？不会修改 Excel。") == QMessageBox.StandardButton.Yes:
            try:
                self.production_backend.preferences()
                field_mapper.remove_field_mapping(str(self.production_backend.config_dir), list(mappings)[entries.index(selected)])
                self.invalidate()
            except (ValueError, OSError) as exc:
                QMessageBox.warning(self, "映射未移除", str(exc))

    def show_page(self, page):
        super().show_page(page)
        if page == "history" and getattr(self, "_configured", False) and not self.busy:
            self.refresh_history()

    def build_feedback(self):
        page = self.simple_page("问题反馈与功能建议", [
            "遇到问题或有改进建议，可以自愿填写腾讯问卷。",
            "在线反馈 · 点击下方按钮并确认后，将在系统浏览器打开第三方问卷，需要联网。",
            "离线业务 · 汇总、对账和导出仍在本地运行；不开启反馈不影响使用。",
            "隐私保护 · 不会自动附带账单、日志、机器码或授权信息。请勿在问卷填写密码、私钥或真实业务数据；截图请先遮盖私人信息。",
            "反馈内容 · 建议说明软件版本、所用功能、操作步骤及错误提示。联系方式自愿填写。",
            "服务说明 · 问卷不是实时客服；填写内容由腾讯问卷平台收集。"])
        self.feedback_button = push("打开在线反馈 ↗", "Link")
        self.feedback_button.clicked.connect(self.open_online_feedback)
        page.layout().insertWidget(page.layout().count() - 1, self.feedback_button)
        return page

    def confirm_online_feedback(self):
        dialog = QMessageBox(self)
        dialog.setWindowTitle("打开在线反馈")
        dialog.setIcon(QMessageBox.Icon.Question)
        dialog.setText(
            "即将在系统浏览器打开腾讯问卷，需要联网。\n"
            "不会自动上传文件或附带机器码、授权信息。是否继续？")
        open_button = dialog.addButton("打开问卷", QMessageBox.ButtonRole.AcceptRole)
        cancel_button = dialog.addButton("取消", QMessageBox.ButtonRole.RejectRole)
        dialog.setDefaultButton(cancel_button)
        dialog.setEscapeButton(cancel_button)
        dialog.exec()
        confirmed = dialog.clickedButton() is open_button
        dialog.deleteLater()
        return confirmed

    def open_online_feedback(self):
        url = "https://wj.qq.com/s2/28242546/gc6x/"
        if not self.confirm_online_feedback():
            return
        try:
            opened = QDesktopServices.openUrl(QUrl(url))
        except Exception:
            opened = False
        if not opened:
            QMessageBox.information(self, "未能打开浏览器",
                "请手动复制以下地址到浏览器打开：\n" + url
                + "\n离线汇总、对账和导出不受影响。")

    def build_about(self):
        return self.simple_page("关于 E-help", ["专注于 Excel 汇总、月度处理与账单对账。",
            "数据与隐私 · 业务处理离线运行，账单和结果保存在本机；在线反馈为自愿使用的外部网页。",
            "主要功能 · 常规汇总、月度汇总与账单对账。",
            "第三方许可 · 图标与依赖的许可说明随软件提供。"])

    def build_settings(self):
        page = self.simple_page("软件版本", [APP_FULL_NAME,
            "汇总、对账与导出在本机完成，不上传账单；在线反馈仅由你主动打开。",
            "使用帮助可在使用教程中查看。"])
        layout = page.layout()
        self.edition_button = push("版本状态", "Link")
        self.edition_button.clicked.connect(self.show_edition_status)
        layout.insertWidget(layout.count() - 1, self.edition_button)
        return page

    def show_edition_status(self):
        message = edition_status()
        dialog = QMessageBox(self)
        dialog.setWindowTitle("版本状态")
        dialog.setText(message)
        dialog.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        dialog.setStandardButtons(QMessageBox.StandardButton.Close)
        dialog.button(QMessageBox.StandardButton.Close).setText("关闭")
        dialog.exec()

    def closeEvent(self, event):
        if self.busy:
            if self._close_pending:
                event.ignore()
                return
            message = "常规任务需等待完成后关闭。是否等待并退出？" if self._job_request.mode == "regular" else "取消当前任务，等待资源释放后退出？已完成保存的文件会保留。"
            if QMessageBox.question(self, "任务运行中", message) == QMessageBox.StandardButton.Yes:
                self._close_pending = True
                self.cancel_job()
                self.preview_button.setEnabled(False)
                self.export_button.setEnabled(False)
            event.ignore()
            return
        self.poll_timer.stop()
        super().closeEvent(event)
