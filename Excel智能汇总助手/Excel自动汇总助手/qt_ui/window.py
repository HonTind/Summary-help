"""Native, interactive, offline Qt preview. Never opens real user data."""
from html import escape
from PySide6.QtCore import Qt, QSize, QTimer, QEvent
from PySide6.QtWidgets import (QApplication, QButtonGroup, QCheckBox, QComboBox,
    QDialog, QDialogButtonBox, QFrame, QHBoxLayout, QInputDialog, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QMainWindow, QMessageBox, QScrollArea,
    QStackedWidget, QTextEdit, QVBoxLayout, QWidget)
from ..offline_help import ARTICLES, search_articles
from ..config import APP_NAME
from .assets import Icons
from .demo import DemoBackend, PreviewRequest, TableData
from .widgets import Disclosure, TablePanel, push, text


class PreviewWindow(QMainWindow):
    MODES = {"regular": "常规汇总", "monthly": "月度汇总", "compare": "账单对账"}
    PAGES = {"workbench": "工作台", "history": "任务记录", "rules": "规则与模板",
             "help": "使用教程", "feedback": "问题反馈", "settings": "设置"}

    def __init__(self, backend=None, *, available_size=None, populate_demo=True):
        super().__init__()
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.backend = backend or DemoBackend()
        self.icons = Icons()
        self._image_labels = []
        self.mode, self.page, self.busy = "monthly", "workbench", False
        self.result = None
        self.selected_files = list(DemoBackend.samples)
        self.memory_templates = {"标准月度规则": ("客户", "金额")}
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.finish_active_preview)
        self.active_request = None
        self.controls, self.nav_buttons, self.mode_buttons = [], {}, {}
        self.setWindowTitle("E助手 · Qt 界面预览（临时示例）")
        self.setWindowIcon(self.icons.brand())
        screen = QSize(*available_size) if available_size else self.screen().availableGeometry().size()
        self.resize(min(1168, screen.width() - 40), min(824, screen.height() - 60))
        self.setMinimumSize(min(840, screen.width() - 40), min(580, screen.height() - 60))
        root = QWidget()
        root.setObjectName("Root")
        self.setCentralWidget(root)
        layout = QHBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        layout.addWidget(self.build_sidebar())
        main = QWidget()
        main.setObjectName("Main")
        main_layout = QVBoxLayout(main)
        main_layout.setContentsMargins(14, 14, 14, 0)
        main_layout.setSpacing(0)
        heading = QHBoxLayout()
        heading.setContentsMargins(18, 0, 8, 0)
        self.title = text("月度汇总", "Title")
        heading.addWidget(self.title)
        heading.addStretch()
        self.help_button = push("查看本功能教程", "Link", icon=self.icons.get("external", active=True, size=18))
        self.help_button.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.help_button.setToolTip("打开软件自带的离线教程。")
        self.help_button.clicked.connect(self.context_help)
        heading.addWidget(self.help_button)
        main_layout.addLayout(heading)
        main_layout.addSpacing(8)
        self.mode_bar = QFrame()
        self.mode_bar.setObjectName("TabBar")
        modes = QHBoxLayout(self.mode_bar)
        modes.setContentsMargins(8, 0, 0, 0)
        modes.setSpacing(16)
        self.mode_group = QButtonGroup(self)
        for mode, caption in self.MODES.items():
            button = push(caption, "Mode")
            button.setCheckable(True)
            button.setFixedHeight(40)
            button.clicked.connect(lambda _checked=False, value=mode: self.select_mode(value))
            self.mode_group.addButton(button)
            self.mode_buttons[mode] = button
            modes.addWidget(button)
        modes.addStretch()
        main_layout.addWidget(self.mode_bar)
        main_layout.addSpacing(10)
        self.pages = QStackedWidget()
        self.page_widgets = {}
        self.page_widgets["workbench"] = self.build_workbench()
        self.page_widgets["history"] = self.build_history()
        self.page_widgets["rules"] = self.build_rules()
        self.page_widgets["help"] = self.build_help()
        self.page_widgets["feedback"] = self.build_feedback()
        self.page_widgets["settings"] = self.build_settings()
        settings_layout = self.page_widgets["settings"].layout()
        settings_layout.insertWidget(settings_layout.count() - 1, self.build_about())
        for page in self.page_widgets.values():
            self.pages.addWidget(page)
        main_layout.addWidget(self.pages, 1)
        main_layout.addSpacing(10)
        self.footer = self.build_footer()
        main_layout.addWidget(self.footer)
        layout.addWidget(main, 1)
        self.select_mode("monthly")
        if populate_demo:
            self.show_result(self.backend.preview(self.request()))
            self.export_button.setEnabled(True)
            self.export_button.setToolTip("已载入演示预览；本按钮只说明后续导出接入，不写任何文件。")

    def build_sidebar(self):
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(190)
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(10, 22, 10, 26)
        layout.setSpacing(4)
        brand = QHBoxLayout()
        brand.setContentsMargins(8, 0, 0, 0)
        icon = self.image_label("brand", size=30)
        icon.setFixedSize(30, 30)
        brand.addWidget(icon)
        brand.addWidget(text(APP_NAME, "Brand"))
        brand.addStretch()
        layout.addLayout(brand)
        layout.addSpacing(12)
        self.nav_group = QButtonGroup(self)
        for page, caption in self.PAGES.items():
            if page == "history":
                layout.addStretch(1)
            button = push("  " + caption, "Nav", icon=self.icons.get(page))
            button.setCheckable(True)
            button.setFixedHeight(46)
            button.clicked.connect(lambda _checked=False, value=page: self.show_page(value))
            self.nav_group.addButton(button)
            self.nav_buttons[page] = button
            layout.addWidget(button)
        return sidebar

    def image_label(self, name, *, size=22, active=False):
        label = QLabel()
        icon = self.icons.brand() if name == "brand" else self.icons.get(name, active=active, size=size)
        label.setPixmap(icon.pixmap(QSize(size, size), self.devicePixelRatioF()))
        self._image_labels.append((label, icon, size))
        label.destroyed.connect(lambda _object=None, reference=label: self.forget_image(reference))
        return label

    def forget_image(self, label):
        self._image_labels = [entry for entry in self._image_labels if entry[0] is not label]

    def event(self, event):
        if event.type() == QEvent.Type.DevicePixelRatioChange and hasattr(self, "_image_labels"):
            for label, icon, size in self._image_labels:
                if label.parent() is not None:
                    label.setPixmap(icon.pixmap(QSize(size, size), self.devicePixelRatioF()))
        return super().event(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "footer_hint"):
            self.footer_hint.setText("仅演示 · 无文件写入" if self.width() < 1000
                                     else "实际使用需先预览；导出不覆盖原文件。")

    def card(self):
        frame = QFrame()
        frame.setObjectName("Card")
        return frame

    def field(self, layout, caption, values, *, editable=True, gap=18):
        layout.addSpacing(gap)
        heading = text(caption, "FieldLabel")
        layout.addWidget(heading)
        layout.addSpacing(6)
        control = QComboBox()
        control.field_label = heading
        control.addItems(values)
        control.setEditable(editable)
        control.setFixedHeight(36)
        control.currentTextChanged.connect(self.invalidate)
        layout.addWidget(control)
        self.controls.append(control)
        return control

    def build_workbench(self):
        page = QWidget()
        row = QHBoxLayout(page)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(12)
        settings = self.card()
        settings.setFixedWidth(246)
        shell = QVBoxLayout(settings)
        shell.setContentsMargins(1, 1, 1, 1)
        scroll = QScrollArea()
        scroll.setObjectName("Settings")
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        body = QWidget()
        body.setObjectName("CardBody")
        self.settings_scroll = scroll
        form = QVBoxLayout(body)
        self.settings_form = form
        form.setContentsMargins(16, 16, 16, 16)
        form.setSpacing(0)
        caption = QHBoxLayout()
        caption.setSpacing(14)
        caption.addWidget(text("本次输入", "FieldLabel"))
        self.file_count = text("2个文件", "Muted")
        caption.addWidget(self.file_count)
        caption.addStretch()
        form.addLayout(caption)
        form.addSpacing(12)
        self.file_card = QFrame()
        self.file_card.setObjectName("Files")
        self.file_layout = QVBoxLayout(self.file_card)
        self.file_layout.setContentsMargins(9, 6, 9, 6)
        self.file_layout.setSpacing(0)
        form.addWidget(self.file_card)
        self.refresh_files()
        form.addSpacing(8)
        actions = QHBoxLayout()
        choose = push("更换文件", "Link")
        self.choose_files_button = choose
        choose.setObjectName("ChooseFiles")
        choose.clicked.connect(self.choose_samples)
        choose.setToolTip("本样板只选择内置演示文件，不浏览你的真实账单。")
        actions.addWidget(choose)
        actions.addStretch()
        form.addLayout(actions)
        self.controls.append(choose)
        self.template = self.field(form, "模板", ["标准月度规则"], editable=True, gap=26)
        self.template.activated.connect(self.load_memory_template)
        self.month = self.field(form, "月份", ["2026-10", "2026-09", "2026-11"])
        self.month.lineEdit().addAction(self.icons.get("calendar", size=18), QLineEdit.ActionPosition.LeadingPosition)
        self.key = self.field(form, "键字段", ["客户", "单据号", "#1"])
        self.amount = self.field(form, "金额字段", ["金额", "#2"])
        form.addSpacing(18)
        self.date_check = QCheckBox("按日期筛选")
        self.date_check.toggled.connect(self.date_toggled)
        form.addWidget(self.date_check)
        self.controls.append(self.date_check)
        self.date_field = QLineEdit("日期")
        self.date_field.setPlaceholderText("日期字段名称或 #列号")
        self.date_field.textChanged.connect(self.invalidate)
        self.date_field.setVisible(False)
        form.addWidget(self.date_field)
        self.controls.append(self.date_field)
        form.addSpacing(6)
        self.date_hint = text("关闭时，月份仅作标签。", "Muted", wrap=True)
        form.addWidget(self.date_hint)
        self.compare_selector = QComboBox()
        self.compare_selector.addItems(["客户总额", "单据逐笔"])
        self.compare_selector.currentTextChanged.connect(self.invalidate)
        self.compare_selector.setVisible(False)
        form.addWidget(self.compare_selector)
        self.controls.append(self.compare_selector)
        form.addSpacing(16)
        advanced = Disclosure("高级设置", self.icons)
        self.advanced = advanced
        advanced_form = QVBoxLayout(advanced.body)
        self.advanced_form = advanced_form
        advanced_form.setContentsMargins(0, 0, 0, 0)
        advanced_form.addWidget(text("演示规则只保存在本次内存。", "Small", wrap=True))
        self.large_sample = QCheckBox("启用大样例（分页）")
        self.large_sample.setToolTip("额外生成201条合成客户用于翻页；不读写真实账单。")
        self.large_sample.toggled.connect(self.invalidate)
        advanced_form.addWidget(self.large_sample)
        self.controls.append(self.large_sample)
        self.save_template = push("保存演示模板", "Link")
        self.save_template.clicked.connect(self.save_memory_template)
        advanced_form.addWidget(self.save_template)
        self.controls.append(self.save_template)
        form.addWidget(advanced)
        form.addStretch()
        self.form_spacers = [(form.itemAt(i).spacerItem(), form.itemAt(i).spacerItem().sizeHint().height())
                             for i in range(form.count()) if form.itemAt(i).spacerItem()
                             and form.itemAt(i).spacerItem().sizeHint().height() > 0]
        scroll.setWidget(body)
        shell.addWidget(scroll)
        row.addWidget(settings)
        right = self.card()
        result_layout = QVBoxLayout(right)
        result_layout.setContentsMargins(16, 16, 16, 16)
        result_layout.setSpacing(0)
        result_layout.addWidget(text("结果预览", "CardTitle"))
        result_layout.addSpacing(12)
        tabs = QFrame()
        tabs.setObjectName("TabBar")
        tab_row = QHBoxLayout(tabs)
        tab_row.setContentsMargins(0, 0, 0, 0)
        tab_row.setSpacing(10)
        self.result_group = QButtonGroup(self)
        self.result_tabs = []
        self.result_stack = QStackedWidget()
        self.table_panels = [TablePanel(self.icons) for _ in range(3)]
        for index, caption in enumerate(("汇总结果", "来源明细", "异常清单")):
            button = push(caption, "ResultTab")
            button.setCheckable(True)
            button.setFixedHeight(40)
            button.clicked.connect(lambda _checked=False, value=index: self.result_stack.setCurrentIndex(value))
            self.result_group.addButton(button)
            self.result_tabs.append(button)
            tab_row.addWidget(button)
            self.result_stack.addWidget(self.table_panels[index])
        self.result_tabs[0].setChecked(True)
        tab_row.addStretch()
        result_layout.addWidget(tabs)
        result_layout.addSpacing(12)
        info = QFrame()
        info.setObjectName("Info")
        info_row = QHBoxLayout(info)
        info_row.setContentsMargins(12, 9, 12, 9)
        icon = self.image_label("about", active=True, size=20)
        info_row.addWidget(icon)
        self.info = text("演示数据，仅用于展示界面。", wrap=True)
        info_row.addWidget(self.info, 1)
        result_layout.addWidget(info)
        result_layout.addSpacing(10)
        result_layout.addWidget(self.result_stack, 1)
        result_layout.addSpacing(24)
        separator = QFrame()
        separator.setObjectName("Separator")
        result_layout.addWidget(separator)
        result_layout.addSpacing(10)
        details = Disclosure("运行详情（点击展开）", self.icons)
        details_layout = QVBoxLayout(details.body)
        details_layout.setContentsMargins(0, 0, 0, 0)
        self.status = text("本地界面样板，只有内置演示数据；未接入真实账单与导出。", "Small", wrap=True)
        details_layout.addWidget(self.status)
        result_layout.addWidget(details)
        row.addWidget(right, 1)
        return page

    def build_footer(self):
        footer = QFrame()
        footer.setObjectName("Footer")
        layout = QHBoxLayout(footer)
        layout.setContentsMargins(16, 10, 0, 12)
        icon = self.image_label("about", size=16)
        layout.addWidget(icon)
        self.footer_hint = text("实际使用需先预览；导出不覆盖原文件。", "Small", wrap=True)
        layout.addWidget(self.footer_hint, 1)
        self.preview_button = push("重新预览")
        self.preview_button.setFixedSize(120, 42)
        self.preview_button.clicked.connect(self.preview)
        self.export_button = push("导出结果", "Primary")
        self.export_button.setFixedSize(132, 42)
        self.export_button.clicked.connect(self.explain_export)
        layout.addWidget(self.preview_button)
        layout.addWidget(self.export_button)
        return footer

    def refresh_files(self):
        while self.file_layout.count():
            item = self.file_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for name in self.selected_files:
            row = QWidget()
            layout = QHBoxLayout(row)
            layout.setContentsMargins(0, 4, 0, 4)
            layout.setSpacing(10)
            icon = self.image_label("file")
            layout.addWidget(icon)
            filename = text(name)
            filename.setToolTip("内置演示文件名，不代表读取了此路径。")
            layout.addWidget(filename, 1)
            self.file_layout.addWidget(row)
        if not self.selected_files:
            self.file_layout.addWidget(text("尚未选择演示文件", "Small"))
        self.file_count.setText(f"{len(self.selected_files)}个文件")

    def choose_samples(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("选择演示输入 · 不读取真实账单")
        layout = QVBoxLayout(dialog)
        layout.addWidget(text("只选择内置样例，不会浏览或打开本地业务文件。", "Small", wrap=True))
        choices = QListWidget()
        for name in DemoBackend.samples:
            item = QListWidgetItem(name)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if name in self.selected_files else Qt.CheckState.Unchecked)
            choices.addItem(item)
        layout.addWidget(choices)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.selected_files = [choices.item(index).text() for index in range(choices.count())
                                   if choices.item(index).checkState() == Qt.CheckState.Checked]
            self.refresh_files()
            self.invalidate()

    def request(self):
        return PreviewRequest(self.mode, tuple(self.selected_files), self.month.currentText(),
            self.key.currentText(), self.amount.currentText(), self.date_check.isChecked(),
            self.date_field.text(), self.compare_selector.currentText(), 201 if self.large_sample.isChecked() else 0)

    def invalidate(self, *_):
        if hasattr(self, "export_button"):
            self.result = None
            self.export_button.setEnabled(False)
            self.info.setText("输入或规则已更改，请重新预览。演示不处理真实账单。")
            for panel in self.table_panels:
                panel.set_data(TableData(panel.data.headers, ()))

    def date_toggled(self, checked):
        self.date_field.setVisible(checked)
        self.date_hint.setText("开启时，只演示所选月份的筛选。" if checked else "关闭时，月份仅作标签。")
        self.invalidate()

    def preview(self):
        if self.busy:
            return
        request = self.request()
        self.busy = True
        for control in (*self.controls, *self.mode_buttons.values(), self.preview_button, self.export_button):
            control.setEnabled(False)
        for key in ("workbench", "history", "rules"):
            self.nav_buttons[key].setEnabled(False)
        self.info.setText("正在预览内置演示数据…")
        self.active_request = request
        self.timer.start(160)

    def finish_active_preview(self):
        request, self.active_request = self.active_request, None
        if request is not None:
            self.finish_preview(request)

    def finish_preview(self, request):
        try:
            self.show_result(self.backend.preview(request))
        except ValueError as error:
            self.result = None
            self.info.setText(str(error))
        finally:
            self.busy = False
            for control in (*self.controls, *self.mode_buttons.values(), self.preview_button):
                control.setEnabled(True)
            for key in ("workbench", "history", "rules"):
                self.nav_buttons[key].setEnabled(True)
            self.export_button.setEnabled(self.result is not None)

    def show_result(self, result):
        self.result = result
        for panel, data in zip(self.table_panels, (result.summary, result.sources, result.issues)):
            panel.set_data(data)
        self.info.setText(result.message)
        self.status.setText(f"演示预览完成 · {len(result.summary.rows)}条结果 · {len(result.sources.rows)}条来源。没有写入真实任务记录。")

    def explain_export(self):
        QMessageBox.information(self, "界面样板 · 尚未接入真实导出",
            "这个阶段用于验收界面。\n不会生成、覆盖或上传任何Excel文件。\n正式导出将在视觉验收后接入现有业务引擎。")

    def save_memory_template(self):
        name = self.template.currentText().strip()
        if not name:
            QMessageBox.warning(self, "演示模板", "请填写模板名称。")
            return
        fields = (self.key.currentText(), self.amount.currentText())
        if name in self.memory_templates and self.memory_templates[name] != fields:
            answer = QMessageBox.question(self, "更新内存演示模板", "仅更新本次演示模板？不会修改真实配置。")
            if answer != QMessageBox.StandardButton.Yes:
                return
        self.memory_templates[name] = fields
        if self.template.findText(name) < 0:
            self.template.addItem(name)
        self.info.setText("演示模板已保存在本次内存；关闭窗口即丢弃，不写真实配置。")

    def load_memory_template(self, *_):
        fields = self.memory_templates.get(self.template.currentText())
        if fields:
            self.key.setCurrentText(fields[0])
            self.amount.setCurrentText(fields[1])

    def select_mode(self, mode):
        if self.busy or mode not in self.MODES:
            return
        self.mode = mode
        self.mode_buttons[mode].setChecked(True)
        self.month.setVisible(mode == "monthly")
        self.month.field_label.setVisible(mode == "monthly")
        self.date_check.setVisible(mode == "monthly")
        self.date_field.setVisible(mode == "monthly" and self.date_check.isChecked())
        self.date_hint.setVisible(mode == "monthly")
        self.compare_selector.setVisible(mode == "compare")
        self.show_page("workbench")
        self.invalidate()

    def show_page(self, page):
        if page not in self.PAGES or (self.busy and page in ("workbench", "rules", "history")):
            return
        self.page = page
        self.pages.setCurrentWidget(self.page_widgets[page])
        self.title.setText(self.MODES[self.mode] if page == "workbench" else self.PAGES[page])
        workbench = page == "workbench"
        self.mode_bar.setVisible(workbench)
        self.footer.setVisible(workbench)
        self.help_button.setVisible(workbench)
        for key, button in self.nav_buttons.items():
            button.setChecked(key == page)
            button.setIcon(self.icons.get(key, active=key == page))

    def context_help(self):
        self.show_page("help")
        self.help_search.clear()
        match = next((index for index, article in enumerate(self.article_list) if article.mode == self.mode and article.title != "快速开始"), 0)
        self.help_topics.setCurrentRow(match)

    def simple_page(self, heading, lines):
        card = self.card()
        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)
        layout.addWidget(text(heading, "CardTitle"))
        for index, line in enumerate(lines):
            paragraph = text(line, "PageLead" if index == 0 else "PageBody", wrap=True)
            paragraph.setMaximumWidth(760)
            paragraph.setTextFormat(Qt.TextFormat.PlainText)
            paragraph.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            layout.addWidget(paragraph)
        layout.addStretch()
        return card

    def build_settings(self):
        return self.simple_page("设置", ["所有功能均在本机运行，无联网服务。"])

    def build_history(self):
        return self.simple_page("本地任务记录", ["本样板不读取或写入你的真实任务记录。",
            "下一阶段继续复用原历史格式、原数据目录及500条上限，不迁移或清空已有记录。"])

    def build_rules(self):
        card = self.simple_page("规则与模板", ["第一阶段仅有内存演示模板。真实字段规则与模板文件未修改。"])
        button = push("设置演示月度规则", "Link")
        button.clicked.connect(lambda: self.select_mode("monthly"))
        card.layout().insertWidget(2, button, alignment=Qt.AlignmentFlag.AlignLeft)
        return card

    def build_help(self):
        card = self.card()
        layout = QHBoxLayout(card)
        layout.setContentsMargins(20, 20, 20, 20)
        left = QVBoxLayout()
        self.help_search = QLineEdit()
        self.help_search.setPlaceholderText("搜索离线教程")
        left.addWidget(self.help_search)
        self.help_topics = QListWidget()
        self.help_topics.setFixedWidth(210)
        left.addWidget(self.help_topics)
        layout.addLayout(left)
        right = QVBoxLayout()
        self.help_text = QTextEdit()
        self.help_text.setReadOnly(True)
        right.addWidget(self.help_text, 1)
        jump = push("前往相关功能", "Link")
        jump.clicked.connect(self.jump_help)
        right.addWidget(jump, alignment=Qt.AlignmentFlag.AlignLeft)
        layout.addLayout(right, 1)
        self.help_topics.currentRowChanged.connect(self.show_article)
        self.help_search.textChanged.connect(self.refresh_articles)
        self.refresh_articles()
        return card

    def refresh_articles(self, *_):
        self.article_list = search_articles(self.help_search.text())
        self.help_topics.clear()
        self.help_topics.addItems([article.title for article in self.article_list])
        self.help_topics.setCurrentRow(0)

    def show_article(self, index):
        if not 0 <= index < len(self.article_list):
            self.help_text.setPlainText("没有匹配的离线教程。")
            return
        article = self.article_list[index]
        paragraphs = "".join("<p>" + escape(line) + "</p>"
                             for line in article.body.splitlines() if line.strip())
        self.help_text.setHtml(
            "<html><head><style>body {color:#344054; font-size:16px;}"
            "h2 {color:#1D1D1F; font-size:24px; margin-bottom:20px;}"
            "p {margin-top:0; margin-bottom:14px; line-height:150%;}"
            "</style></head><body><h2>" + escape(article.title) + "</h2>"
            + paragraphs + "</body></html>")

    def jump_help(self):
        index = self.help_topics.currentRow()
        if 0 <= index < len(self.article_list):
            self.select_mode(self.article_list[index].mode)

    def build_feedback(self):
        return self.simple_page("联系与反馈", ["遇到问题时，请记录操作步骤与提示内容。",
            "联系方式 · 暂未提供",
            "反馈内容 · 软件版本、所用功能、复现步骤与完整错误提示。",
            "隐私保护 · 截图请遮盖账单、个人信息、文件路径和授权信息；不要发送密码、私钥、API配置或真实业务文件。",
            "离线说明 · 本页不提交、不上传，也不会自动发送任何内容。"])

    def build_about(self):
        return self.simple_page("关于 E助手", ["PySide6本地界面预览 · 第一阶段", "全部演示在本机完成，无账号、云端、遥测或联网服务。",
            "当前正式Tk程序与现有业务引擎保留；视觉验收后逐步接入。", "图标：Feather 4.29.2（MIT）与现有软件logo。"])
