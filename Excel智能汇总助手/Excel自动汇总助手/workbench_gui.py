"""Monthly/reconciliation workbench. Threads never access Tk or legacy settings."""
from __future__ import annotations

import datetime
import json
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .analysis import CancelledError, SheetRules, check_cancel, date_trace_text, parse_amount
from .monthly import MonthlyResult, load_templates, make_template, run_monthly, save_template
from .reconciliation import CATEGORIES, run_reconciliation
from .workbench_export import export_result

BG, SURFACE, TEXT, SECONDARY, ACCENT = "#F5F5F7", "#FFFFFF", "#1D1D1F", "#6E6E73", "#0071E3"


# Keep the original import path available to callers.
from .job_runner import JobRunner
from .workbench_view import build_preview, history_counts


class PagedTree(ttk.Frame):
    """Visible page size only; the complete result remains available for export."""
    PAGE_SIZE = 100

    def __init__(self, parent, headers):
        super().__init__(parent, style="WB.TFrame")
        self.rows, self.page = [], 0
        columns = [f"c{i}" for i in range(len(headers))]
        self.tree = ttk.Treeview(self, columns=columns, show="headings", height=8, style="WB.Treeview")
        for column, header in zip(columns, headers):
            self.tree.heading(column, text=header)
            self.tree.column(column, width=150, minwidth=90, stretch=True)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ttk.Scrollbar(self, orient="vertical", command=self.tree.yview).grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=horizontal.set)
        vertical = self.grid_slaves(row=0, column=1)[0]
        self.tree.configure(yscrollcommand=vertical.set)
        footer = ttk.Frame(self, style="WB.TFrame")
        footer.grid(row=2, column=0, sticky="ew", pady=(5, 0))
        self.previous = ttk.Button(footer, text="上一页", command=lambda: self.move(-1))
        self.previous.pack(side="left")
        self.label = ttk.Label(footer, text="0 行", style="WB.Secondary.TLabel")
        self.label.pack(side="left", padx=12)
        self.next = ttk.Button(footer, text="下一页", command=lambda: self.move(1))
        self.next.pack(side="left")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.render()

    def set_rows(self, rows):
        self.rows, self.page = list(rows), 0
        self.render()

    def move(self, amount):
        last = max(0, (len(self.rows) - 1) // self.PAGE_SIZE)
        self.page = max(0, min(last, self.page + amount))
        self.render()

    def render(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        start = self.page * self.PAGE_SIZE
        for row in self.rows[start:start + self.PAGE_SIZE]:
            self.tree.insert("", "end", values=row)
        end = min(start + self.PAGE_SIZE, len(self.rows))
        self.label.configure(text=f"{start + 1 if self.rows else 0}–{end} / {len(self.rows)} 行（完整导出）")
        self.previous.configure(state="normal" if self.page else "disabled")
        self.next.configure(state="normal" if end < len(self.rows) else "disabled")


def column_selection(text):
    text = text.strip()
    if text.startswith("#"):
        if not text[1:].isdigit():
            raise ValueError("列序号请用 #1、#2 等格式")
        return int(text[1:])
    return text


def column_text(value):
    return f"#{value}" if type(value) is int else value


class InputPanel(ttk.Frame):
    def __init__(self, parent, title, changed):
        super().__init__(parent, padding=12, style="WB.Card.TFrame")
        self.paths = []
        self.changed = changed
        self.controls = []
        self.sheet = tk.StringVar()
        self.header = tk.StringVar(value="1")
        self.key = tk.StringVar(value="客户")
        self.amount = tk.StringVar(value="金额")
        self.extra_keys = tk.StringVar()
        bar = ttk.Frame(self, style="WB.Card.TFrame")
        bar.grid(row=0, column=0, columnspan=4, sticky="ew")
        ttk.Label(bar, text=title, style="WB.Card.TLabel").pack(side="left")
        for text, command in (("选择 Excel", self.select_files), ("清空", lambda: self.set_files([]))):
            button = ttk.Button(bar, text=text, command=command)
            button.pack(side="right", padx=(6, 0))
            self.controls.append(button)
        self.files = tk.Listbox(self, height=3, bg="#FAFAFB", fg=TEXT, relief="flat",
                                font=("Microsoft YaHei UI", 9), exportselection=False)
        self.files.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(7, 8))
        self.controls.append(self.files)
        for row, column, label, variable in (
            (2, 0, "工作表", self.sheet), (2, 2, "表头行", self.header),
            (3, 0, "键字段①", self.key), (3, 2, "金额字段", self.amount),
        ):
            ttk.Label(self, text=label, style="WB.Card.TLabel").grid(row=row, column=column, sticky="w", padx=(0, 5))
            entry = ttk.Entry(self, textvariable=variable, width=14)
            entry.grid(row=row, column=column + 1, sticky="ew", pady=3, padx=(0, 8))
            self.controls.append(entry)
            if variable is self.key:
                self.key_entry = entry
            variable.trace_add("write", lambda *_: self.changed())
        ttk.Label(self, text="组合键②③…", style="WB.Card.TLabel").grid(
            row=4, column=0, sticky="w", padx=(0, 5), pady=3)
        self.extra_entry = ttk.Entry(self, textvariable=self.extra_keys, width=14)
        self.extra_entry.grid(row=4, column=1, columnspan=3, sticky="ew", pady=3, padx=(0, 8))
        self.controls.append(self.extra_entry)
        self.extra_keys.trace_add("write", lambda *_: self.changed())
        ttk.Label(self, text="工作表留空用默认表；首键填上方，其余按顺序填 JSON 数组如 [\"单据号\",\"批次\"]；"
                            "字段填精确表头名或 #列号，共 1–8 项。每次重新选择账单。",
                  style="WB.Secondary.TLabel", wraplength=440).grid(row=5, column=0, columnspan=4, sticky="w", pady=(5, 0))
        self.columnconfigure(1, weight=1)
        self.columnconfigure(3, weight=1)

    def select_files(self):
        files = filedialog.askopenfilenames(parent=self, title="选择本次账单（可多选）",
                                           filetypes=[("Excel", "*.xlsx *.xlsm *.xls")])
        if files:
            self.set_files(files)

    def set_files(self, files):
        self.paths = list(dict.fromkeys(str(Path(p).resolve()) for p in files))
        self.files.delete(0, "end")
        for path in self.paths:
            self.files.insert("end", path)
        self.changed()

    def key_fields(self):
        """Ordered key selections: the first field plus an optional JSON array of extras."""
        fields = [column_selection(self.key.get())]
        raw = self.extra_keys.get().strip()
        if raw:
            try:
                data = json.loads(raw)
            except json.JSONDecodeError as exc:
                raise ValueError('组合键后续字段需为 JSON 数组，例如 ["单据号", "批次"]') from exc
            if not isinstance(data, list) or not 1 <= len(data) <= 7:
                raise ValueError("组合键后续字段需为 1 至 7 项的 JSON 数组（加上首键最多 8 项）")
            for item in data:
                if not isinstance(item, str) or not item.strip():
                    raise ValueError('组合键字段需为非空字符串；列序号请写成 "#2"')
                fields.append(column_selection(item))
        return tuple(fields)

    def rules(self):
        try:
            header = int(self.header.get())
        except ValueError as exc:
            raise ValueError("表头行必须是整数") from exc
        fields = self.key_fields()
        if len(fields) == 1:
            return SheetRules(fields[0], column_selection(self.amount.get()),
                              self.sheet.get().strip() or None, header)
        return SheetRules(fields[0], column_selection(self.amount.get()),
                          self.sheet.get().strip() or None, header, fields)

    def apply_rules(self, rules):
        fields = rules.field_selections()
        self.key.set(column_text(fields[0]))
        self.extra_keys.set(json.dumps([column_text(field) for field in fields[1:]], ensure_ascii=False)
                            if len(fields) > 1 else "")
        self.amount.set(column_text(rules.amount_column))
        self.sheet.set(rules.sheet or "")
        self.header.set(str(rules.header_row))


class WorkbenchController:
    """UI-only controller shared by the compatible window and embedded frame."""
    def _init_state(self, template_path, *, history_writer=None):
        self.template_path = Path(template_path)
        self.runner = JobRunner()
        self.result = None
        self._busy = False
        self._closing = False
        self._closed = False
        self._after = None
        self._operation = None
        self._history_writer = history_writer
        self._history_context = None
        self._history_warning = ""
        self._pending = []
        self._suspend = True
        self.bind("<Destroy>", self._destroyed, add="+")
        self._styles()
        self._build()
        self._suspend = False
        self.refresh_templates()

    def _styles(self):
        style = ttk.Style(self)
        style.configure("WB.TFrame", background=BG)
        style.configure("WB.Card.TFrame", background=SURFACE)
        style.configure("WB.TLabel", background=BG, foreground=TEXT, font=("Microsoft YaHei UI", 10))
        style.configure("WB.Card.TLabel", background=SURFACE, foreground=TEXT, font=("Microsoft YaHei UI", 9))
        style.configure("WB.Secondary.TLabel", background=BG, foreground=SECONDARY, font=("Microsoft YaHei UI", 9))
        style.configure("WB.Title.TLabel", background=BG, foreground=TEXT, font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("WB.Treeview", rowheight=26, font=("Microsoft YaHei UI", 9))
        style.configure("WB.Accent.TButton", foreground=ACCENT, padding=(15, 7), font=("Microsoft YaHei UI", 10))

    def _build(self):
        body = ttk.Frame(self, padding=16, style="WB.TFrame")
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(3, weight=1)
        ttk.Label(body, text="月度与对账工作台", style="WB.Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(body, text="复用规则，每次重新选文件。输入只读，金额精确比较。", style="WB.Secondary.TLabel").grid(
            row=1, column=0, sticky="w", pady=(3, 10))
        self.tabs = ttk.Notebook(body)
        self.tabs.grid(row=2, column=0, sticky="ew")
        month_page = ttk.Frame(self.tabs, padding=10, style="WB.TFrame")
        compare_page = ttk.Frame(self.tabs, padding=10, style="WB.TFrame")
        self.tabs.add(month_page, text="每月重复汇总")
        self.tabs.add(compare_page, text="两份账单核对差异")
        month_page.columnconfigure(0, weight=1)
        compare_page.columnconfigure(0, weight=1)
        compare_page.columnconfigure(1, weight=1)
        self.template_name = tk.StringVar(value="默认月报")
        self.month = tk.StringVar(value=datetime.date.today().strftime("%Y-%m"))
        self.output_prefix = tk.StringVar(value="月度汇总")
        self.mode = tk.StringVar(value="客户总额")
        self.tolerance = tk.StringVar(value="0")
        month_bar = ttk.Frame(month_page, style="WB.TFrame")
        month_bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(month_bar, text="模板", style="WB.TLabel").pack(side="left")
        self.template_box = ttk.Combobox(month_bar, textvariable=self.template_name, width=17,
                                        postcommand=self.refresh_templates)
        self.template_box.pack(side="left", padx=6)
        load_button = ttk.Button(month_bar, text="读取模板", command=self.load_template)
        load_button.pack(side="left", padx=3)
        save_button = ttk.Button(month_bar, text="保存规则", command=self.save_template)
        save_button.pack(side="left", padx=3)
        ttk.Label(month_bar, text="月份", style="WB.TLabel").pack(side="left", padx=(16, 3))
        month_entry = ttk.Entry(month_bar, textvariable=self.month, width=9)
        month_entry.pack(side="left")
        ttk.Label(month_bar, text="输出前缀", style="WB.TLabel").pack(side="left", padx=(12, 3))
        prefix_entry = ttk.Entry(month_bar, textvariable=self.output_prefix, width=14)
        prefix_entry.pack(side="left")
        self.date_filter = tk.BooleanVar(value=False)
        self.date_field = tk.StringVar()
        date_bar = ttk.Frame(month_page, style="WB.TFrame")
        date_bar.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        self.date_check = ttk.Checkbutton(date_bar, text="按日期筛选", variable=self.date_filter,
                                          command=self._date_toggle)
        self.date_check.pack(side="left")
        ttk.Label(date_bar, text="日期字段", style="WB.TLabel").pack(side="left", padx=(12, 4))
        self.date_entry = ttk.Entry(date_bar, textvariable=self.date_field, width=16, state="disabled")
        self.date_entry.pack(side="left")
        ttk.Label(date_bar, text="开启后只汇总该月份（YYYY-MM）；字段填精确表头名或 #列号。",
                  style="WB.Secondary.TLabel").pack(side="left", padx=10)
        self.date_field.trace_add("write", lambda *_: self.invalidate())
        self.monthly_input = InputPanel(month_page, "本月账单", self.invalidate)
        self.monthly_input.grid(row=2, column=0, sticky="ew")
        compare_bar = ttk.Frame(compare_page, style="WB.TFrame")
        compare_bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        mode_box = ttk.Combobox(compare_bar, textvariable=self.mode, values=("客户总额", "单据逐笔"),
                               state="readonly", width=12)
        mode_box.pack(side="left")
        mode_box.bind("<<ComboboxSelected>>", self.mode_changed)
        ttk.Label(compare_bar, text="金额容差", style="WB.TLabel").pack(side="left", padx=(16, 6))
        tolerance_entry = ttk.Entry(compare_bar, textvariable=self.tolerance, width=12)
        tolerance_entry.pack(side="left")
        ttk.Label(compare_bar, text="默认 0；单据重复会列为歧义，不自动匹配。", style="WB.Secondary.TLabel").pack(
            side="left", padx=12)
        self.left_input = InputPanel(compare_page, "左侧账单", self.invalidate)
        self.right_input = InputPanel(compare_page, "右侧账单", self.invalidate)
        self.left_input.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
        self.right_input.grid(row=1, column=1, sticky="nsew", padx=(6, 0))
        self.controls = [(widget, str(widget.cget("state"))) for widget in
                         (self.template_box, load_button, save_button, month_entry, prefix_entry,
                          self.date_check, mode_box, tolerance_entry, *self.monthly_input.controls,
                          *self.left_input.controls, *self.right_input.controls)]
        self.views = ttk.Notebook(body)
        self.views.grid(row=3, column=0, sticky="nsew", pady=(12, 6))
        self.summary = PagedTree(self.views, ("键", "分类", "汇总 / 左金额", "右金额", "左减右", "来源行数"))
        self.sources = PagedTree(self.views, ("侧别", "分类", "文件", "工作表", "原行号", "键",
                                              "精确金额", "日期（原｜规范）"))
        self.issues = PagedTree(self.views, ("侧别", "文件", "工作表", "原行号", "异常", "说明",
                                            "原始键", "原始金额", "原日期", "规范日期"))
        for frame, text in ((self.summary, "结果"), (self.sources, "来源明细"), (self.issues, "异常清单")):
            self.views.add(frame, text=text)
        footer = ttk.Frame(body, style="WB.TFrame")
        footer.grid(row=4, column=0, sticky="ew")
        self.preview_button = ttk.Button(footer, text="预览结果", style="WB.Accent.TButton", command=self.preview)
        self.preview_button.pack(side="left")
        self.cancel_button = ttk.Button(footer, text="取消", command=self.cancel, state="disabled")
        self.cancel_button.pack(side="left", padx=8)
        self.export_button = ttk.Button(footer, text="导出完整 Excel", command=self.export, state="disabled")
        self.export_button.pack(side="left")
        self.status = tk.StringVar(value="请选择文件和字段，再预览。默认月份仅作标签；开启按日期筛选后只汇总所选月份。")
        ttk.Label(body, textvariable=self.status, style="WB.Secondary.TLabel", wraplength=950).grid(
            row=5, column=0, sticky="w", pady=(7, 0))
        self._active_tab = self.tabs.select()
        self.tabs.bind("<<NotebookTabChanged>>", self.tab_changed)
        for variable in (self.template_name, self.month, self.output_prefix, self.mode, self.tolerance):
            variable.trace_add("write", lambda *_: self.invalidate())

    def invalidate(self):
        if self._suspend or self._busy or self._closed:
            return
        self.result = None
        self.export_button.configure(state="disabled")
        for view in (self.summary, self.sources, self.issues):
            view.set_rows([])
        self.status.set("输入或规则已变化，请重新预览；旧结果不会用于导出。")

    def _date_toggle(self):
        if not self._busy and not self._closed:
            self.date_entry.configure(state="normal" if self.date_filter.get() else "disabled")
        self.invalidate()

    def monthly_rules(self):
        """Monthly rules, adding the selected date field when the filter is on."""
        rules = self.monthly_input.rules()
        if not self.date_filter.get():
            return rules
        raw = self.date_field.get().strip()
        if not raw:
            raise ValueError("已开启按日期筛选，请填写日期字段（精确表头名或 #列号）")
        return SheetRules(rules.key_column, rules.amount_column, rules.sheet, rules.header_row,
                          rules.key_fields, column_selection(raw))

    def tab_changed(self, _event=None):
        selected = self.tabs.select()
        if not self._busy and selected != self._active_tab:
            self._active_tab = selected
            self.invalidate()

    def mode_changed(self, _event=None):
        default = "单据号" if self.mode.get() == "单据逐笔" else "客户"
        for panel in (self.left_input, self.right_input):
            # Never overwrite a user-configured composite; only nudge the untouched default.
            if not panel.extra_keys.get().strip() and panel.key.get() in ("客户", "单据号"):
                panel.key.set(default)
        self.invalidate()

    def load_template(self):
        if self._busy:
            return
        try:
            templates = load_templates(self.template_path)
            self.template_box.configure(values=tuple(templates))
            name = self.template_name.get().strip()
            if name not in templates:
                raise ValueError("未找到该模板；可从下拉列表选择已存模板，或设置规则后保存新模板")
            template = templates[name]
            self.monthly_input.apply_rules(SheetRules(**template["rules"]))
            self.output_prefix.set(template["output_prefix"])
            date_column = template["rules"].get("date_column")
            self.date_filter.set(date_column is not None)
            self.date_field.set(column_text(date_column) if date_column is not None else "")
            self._date_toggle()
            self.status.set("模板规则已读取。月份和本次文件不会自动更改；请重新预览。")
        except Exception as exc:
            messagebox.showerror("读取模板失败", str(exc), parent=self)

    def refresh_templates(self):
        try:
            self.template_box.configure(values=tuple(load_templates(self.template_path)))
        except Exception as exc:
            self.status.set(f"模板文件无法读取，原文件保留：{exc}")

    def save_template(self):
        if self._busy:
            return
        try:
            template = make_template(self.template_name.get().strip(), self.monthly_rules(),
                                     self.output_prefix.get().strip())
            templates = load_templates(self.template_path)
            allow = template["name"] in templates
            if allow and not messagebox.askyesno("确认更新模板", f"更新「{template['name']}」的规则？不会保存账单内容或文件名单。",
                                                 parent=self):
                return
            save_template(self.template_path, template, allow_update=allow)
            self.template_box.configure(values=tuple(load_templates(self.template_path)))
            self.status.set("模板已保存，只含规则。下个月重新选账单、修改月份即可复用。")
        except Exception as exc:
            messagebox.showerror("保存模板失败", str(exc), parent=self)

    def preview(self):
        if self._busy or self.runner.busy or self._closing:
            return
        try:
            if self.tabs.index(self.tabs.select()) == 0:
                paths = tuple(self.monthly_input.paths)
                if not paths:
                    raise ValueError("请先选择本月 Excel 账单")
                template = make_template(self.template_name.get().strip(), self.monthly_rules(),
                                         self.output_prefix.get().strip())
                month = self.month.get()
                from .monthly import validate_month
                validate_month(month)
                operation = lambda cancel, progress: run_monthly(
                    paths, template, month, cancel=cancel,
                    progress=lambda n, total, name: progress(f"读取 {n}/{total}：{name}"))
            else:
                left, right = tuple(self.left_input.paths), tuple(self.right_input.paths)
                if not left or not right:
                    raise ValueError("请分别选择左侧与右侧 Excel 账单")
                left_rules, right_rules = self.left_input.rules(), self.right_input.rules()
                if len(left_rules.field_selections()) != len(right_rules.field_selections()):
                    raise ValueError("左右两侧键字段数量必须相同，并按顺序一一对应")
                mode = "document" if self.mode.get() == "单据逐笔" else "customer"
                tolerance = parse_amount(self.tolerance.get())
                if tolerance < 0:
                    raise ValueError("容差不能为负数")
                operation = lambda cancel, progress: run_reconciliation(
                    left, right, left_rules, right_rules, mode, tolerance, cancel=cancel,
                    progress=lambda side, n, total, name: progress(f"{side} {n}/{total}：{name}"))
            self.invalidate()
            self._start(operation, "preview")
        except Exception as exc:
            messagebox.showerror("无法开始预览", str(exc), parent=self)

    def _start(self, operation, kind):
        if self._busy or self._closing or self.runner.busy:
            return False
        context = self._history_snapshot(kind)
        # The export module has already passed its final cancellation/save checks.
        # A later cancel must not erase a returned saved path or label it cancelled.
        if not self.runner.start(operation, accept_completed=kind == "export"):
            return False
        self._history_context = context
        self._history_warning = ""
        self._operation = kind
        self._pending.clear()
        self._set_busy(True)
        self.status.set("正在读取与核对…" if kind == "preview" else "正在导出完整结果…")
        self._after = self.after(30, self._poll)
        return True

    def _history_counts(self, result):
        return history_counts(result)

    def _history_snapshot(self, kind):
        if self.tabs.index(self.tabs.select()) == 0:
            task_type, files = "monthly", len(self.monthly_input.paths)
        else:
            task_type = "document" if self.mode.get() == "单据逐笔" else "customer"
            files = len(self.left_input.paths) + len(self.right_input.paths)
        context = {"task_type": task_type, "phase": kind, "file_count": files}
        if kind == "export" and self.result is not None:
            context.update(self._history_counts(self.result))
        return context

    def flush_history(self):
        """Consume terminal metadata once on the Tk thread, also before safe close."""
        if self.runner.busy or self._history_context is None:
            return True
        self._collect()
        context, self._history_context = self._history_context, None
        if self._history_writer is None:
            return True  # Compatible standalone API; opt in with a local writer.
        kind, payload = self._pending[-1] if self._pending else ("error", None)
        if context["phase"] == "preview" and kind == "done" and self.runner.cancel_event.is_set():
            kind = "cancelled"
        status = "cancelled" if kind == "cancelled" else "failed"
        if kind == "done":
            if context["phase"] == "preview":
                context.update(self._history_counts(payload))
                status = "previewed" if payload.complete else "incomplete"
            else:
                status = "success"
                context["output_path"] = payload
        try:
            saved = self._history_writer(status=status, **context)
        except Exception:
            saved = False
        if not saved:
            self._history_warning = "本次记录未保存，原记录文件保留；计算/导出结果不受影响。"
        return bool(saved)

    def _set_busy(self, value):
        if value:
            self._active_tab = self.tabs.select()
        self._busy = value
        for widget, state in self.controls:
            widget.configure(state="disabled" if value else state)
        self.date_entry.configure(state="disabled" if value else ("normal" if self.date_filter.get() else "disabled"))
        for index in range(2):
            self.tabs.tab(index, state="disabled" if value else "normal")
        if not value:
            self.tabs.select(self._active_tab)
        self.preview_button.configure(state="disabled" if value else "normal")
        self.cancel_button.configure(state="normal" if value else "disabled")
        self.export_button.configure(state="normal" if not value and self.result is not None else "disabled")

    def _collect(self):
        while True:
            try:
                kind, payload = self.runner.events.get_nowait()
            except queue.Empty:
                break
            if kind == "progress":
                if not self._closing:
                    self.status.set(payload)
            else:
                self._pending.append((kind, payload))

    def _poll(self):
        self._after = None
        if self._closed:
            return
        self._collect()
        if self.runner.busy:
            self._after = self.after(30, self._poll)
            return
        # Thread has stopped; drain again so a final queue event cannot be lost.
        self._collect()
        saved = self.flush_history()
        if self._closing:
            if not saved:
                messagebox.showwarning("任务记录未保存", self._history_warning, parent=self)
            self.destroy()
            return
        self._set_busy(False)
        if not self._pending:
            self.status.set("任务未返回结果，请重试" + ("；" + self._history_warning if self._history_warning else ""))
            return
        kind, payload = self._pending[-1]
        self._pending.clear()
        if kind == "done" and self._operation == "preview" and self.runner.cancel_event.is_set():
            kind, payload = "cancelled", "任务已取消，预览结果已丢弃"
        if kind == "done":
            if self._operation == "preview":
                self._show_result(payload)
            else:
                self.status.set(f"已导出完整结果：{payload}")
                messagebox.showinfo("导出完成", f"已生成新文件（未覆盖输入或已有结果）：\n{payload}", parent=self)
        else:
            if self._operation == "preview":
                self.result = None
                self.export_button.configure(state="disabled")
            self.status.set(payload)
            if kind == "error":
                messagebox.showerror("任务失败", payload, parent=self)
        if self._history_warning:
            self.status.set(self.status.get() + "；" + self._history_warning)

    def _show_result(self, result):
        self.result = result
        preview = build_preview(result)
        self.summary.set_rows(preview.summary)
        self.sources.set_rows(preview.sources)
        self.issues.set_rows(preview.issues)
        self.views.select(self.summary)
        self.status.set(preview.message)
        self.export_button.configure(state="normal")

    def export(self):
        if self._busy or self.result is None or self._closing:
            return
        if not self.result.complete and not messagebox.askyesno(
            "结果不完整", "存在异常、重复歧义或无有效输入。是否仍导出供核查？文件会保留异常及不完整标记。",
            parent=self):
            return
        directory = filedialog.askdirectory(parent=self, title="选择结果输出目录")
        if directory:
            result = self.result
            self._start(lambda cancel, progress: export_result(result, directory, cancel=cancel), "export")

    def cancel(self):
        if self._busy:
            self.runner.cancel()
            self.status.set("正在取消，等待当前读取／保存结束；部分结果将丢弃。")
            self.cancel_button.configure(state="disabled")

    def close(self):
        if self._closed or self._closing:
            return
        if self._busy or self.runner.busy:
            if not messagebox.askyesno("关闭工作台", "任务仍在运行。取消任务并在安全停止后关闭？", parent=self):
                return
            if self._closed:
                return
            # A native confirmation runs a nested event loop; completion may have
            # been consumed while the question was open.
            if not self.runner.busy:
                if not self.flush_history():
                    messagebox.showwarning("任务记录未保存", self._history_warning, parent=self)
                self.destroy()
                return
            self._closing = True
            self.cancel()
            if self._after is None:
                self._after = self.after(30, self._poll)
        else:
            self.destroy()

    def _destroyed(self, event):
        if event.widget is not self:
            return
        self._closed = True
        self.runner.cancel()
        if self._after is not None:
            try:
                self.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None


class WorkbenchWindow(WorkbenchController, tk.Toplevel):
    """Original standalone entry retained for callers and regression tests."""
    def __init__(self, parent, template_path, *, history_writer=None):
        tk.Toplevel.__init__(self, parent)
        self.title("E助手 · 月度与对账工作台（开发预览）")
        self.configure(bg=BG)
        self.transient(parent)
        self.geometry(f"{max(840, min(1040, self.winfo_screenwidth() - 100))}x"
                      f"{max(580, min(720, self.winfo_screenheight() - 120))}")
        self.minsize(840, 580)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self._init_state(template_path, history_writer=history_writer)
