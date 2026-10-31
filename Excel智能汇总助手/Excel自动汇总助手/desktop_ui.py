"""Offline desktop shell. Reuses the legacy and workbench engines without web services."""
from __future__ import annotations

import datetime
import json
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

from . import gui
from . import history as task_history
from .config import APP_NAME, APP_VERSION, APP_FULL_NAME
from .offline_help import ARTICLES, PUBLIC_CONTACTS, search_articles
from .ui_assets import ICON_PNGS, BRAND_PNGS
from .display import configure_display, configure_styles, window_bounds, px
from .app_icon import apply_app_icon
from .workbench_gui import (ACCENT, BG, SECONDARY, SURFACE, TEXT, InputPanel,
                            PagedTree, WorkbenchController)


def label(parent, text, *, title=False, muted=False, size=None, bold=False, **kwargs):
    item = tk.Label(parent, text=text, bg=SURFACE, fg=SECONDARY if muted else TEXT,
                    font=("Microsoft YaHei UI", size or (24 if title else 12), "bold" if title or bold else "normal"),
                    anchor="w", justify="left", **kwargs)
    if "wraplength" in kwargs:
        maximum = int(kwargs["wraplength"])
        def resize(event):
            if item.winfo_exists() and parent.winfo_width() > 1:
                inset = 2 * sum(int(item.cget(name)) for name in ("padx", "borderwidth", "highlightthickness"))
                available = parent.winfo_width() - 2 * max(0, item.winfo_x()) - inset
                if event.widget is item:
                    available = min(available, event.width - inset)
                item.configure(wraplength=max(1, min(maximum, available)))
        parent.bind("<Configure>", resize, add="+")
        item.bind("<Configure>", resize, add="+")
    return item


def button(parent, text, command, *, primary=False, **kwargs):
    options = dict(font=("Microsoft YaHei UI", 12), bg=ACCENT if primary else SURFACE,
                   fg="white" if primary else TEXT, activebackground="#005BBB" if primary else "#EFF6FF",
                   activeforeground="white" if primary else TEXT, relief="flat", borderwidth=0,
                   padx=px(parent, 12), pady=px(parent, 7), highlightthickness=1, highlightbackground="#DEDEE3", highlightcolor=ACCENT)
    options.update(kwargs)
    control = tk.Button(parent, text=text, command=command, **options)
    control.bind("<Return>", lambda event: control.invoke() if str(control.cget("state")) != "disabled" else None)
    return control


class DesktopTree(PagedTree):
    """Same paging semantics, with the selected design's clearer rows and headings."""
    def __init__(self, parent, headers):
        super().__init__(parent, headers)
        self.tree.configure(style="Desktop.Treeview")
        self.tree.tag_configure("alternate", background="#F7F8FA")
        for column in self.tree["columns"]:
            self.tree.heading(column, anchor="w")
            self.tree.column(column, width=px(self, self.tree.column(column, "width")),
                             minwidth=px(self, self.tree.column(column, "minwidth")))
        footer = self.label.master
        footer.configure(style="WB.Card.TFrame")
        footer.grid_configure(pady=px(self, (10, 0)))
        self.label.configure(style="Desktop.Pager.TLabel")
        vertical = self.grid_slaves(row=0, column=1)[0]
        horizontal = self.grid_slaves(row=1, column=0)[0]
        def scroll_range(bar, first, last):
            bar.set(first, last)
            if float(first) <= 0 and float(last) >= 1:
                bar.grid_remove()
            elif not bar.winfo_manager():
                bar.grid()
        self.tree.configure(yscrollcommand=lambda first, last: scroll_range(vertical, first, last),
                            xscrollcommand=lambda first, last: scroll_range(horizontal, first, last))
        self.previous.pack_forget()
        self.next.pack_forget()
        self.label.pack_forget()
        icons = self.winfo_toplevel()._desktop_icons
        self.previous.configure(text="上一页", image=icons["previous"], compound="left", width=0, style="Desktop.Pager.TButton")
        self.next.configure(text="下一页", image=icons["next"], compound="right", width=0, style="Desktop.Pager.TButton")
        self.next.pack(side="right")
        self.previous.pack(side="right", padx=px(self, (0, 6)))
        self.label.pack(side="left", padx=px(self, 6))
        def fit_pager(event):
            # Keep text labels/commands intact, but show image-only controls
            # when a narrow results card needs room for the complete count.
            compact = event.width < px(self, 420)
            self.previous.configure(compound="none" if compact else "left")
            self.next.configure(compound="none" if compact else "right")
        footer.bind("<Configure>", fit_pager, add="+")

    def render(self):
        super().render()
        for index, item in enumerate(self.tree.get_children()):
            self.tree.item(item, tags=("alternate",) if index % 2 else ())


class HistoryTree(DesktopTree):
    """History paging is not an Excel export feature."""
    def render(self):
        super().render()
        start = self.page * self.PAGE_SIZE
        end = min(start + self.PAGE_SIZE, len(self.rows))
        self.label.configure(text=f"{start + 1 if self.rows else 0}–{end} / {len(self.rows)} 条记录")


class SettingsScroll(tk.Frame):
    """Clipped native frame (not Canvas); keeps long forms accessible on short screens."""
    def __init__(self, parent):
        super().__init__(parent, bg=SURFACE)
        self.inner = tk.Frame(self, bg=SURFACE)
        self.offset = 0
        self.scrollbar = ttk.Scrollbar(self, command=self.scroll)
        self.bind("<Configure>", self.layout)
        self.inner.bind("<Configure>", self.layout)
        self.bind("<MouseWheel>", self.wheel)
        self.inner.bind("<MouseWheel>", self.wheel)

    def layout(self, _event=None):
        if not self.winfo_exists():
            return
        height = max(1, self.inner.winfo_reqheight())
        viewport = max(1, self.winfo_height())
        span = max(0, height - viewport)
        self.offset = min(max(0, self.offset), span)
        bar_width = px(self, 14)
        self.inner.place(x=0, y=-self.offset, width=max(1, self.winfo_width() - (bar_width if span else 0)), height=height)
        if span:
            self.scrollbar.place(relx=1, x=-bar_width, y=0, width=bar_width, relheight=1)
            self.scrollbar.set(self.offset / height, min(1, (self.offset + viewport) / height))
        else:
            self.scrollbar.place_forget()

    def scroll(self, *args):
        span = max(0, self.inner.winfo_reqheight() - self.winfo_height())
        if args[0] == "moveto":
            self.offset = int(float(args[1]) * self.inner.winfo_reqheight())
        elif args[0] == "scroll":
            self.offset += int(args[1]) * (max(px(self, 40), self.winfo_height() - px(self, 40))
                                         if args[2] == "pages" else px(self, 40))
        self.offset = max(0, min(span, self.offset))
        self.layout()

    def wheel(self, event):
        self.scroll("scroll", -1 if event.delta > 0 else 1, "units")
        return "break"

    def bind_form(self):
        pending = [self.inner]
        while pending:
            widget = pending.pop()
            # Lists and entry fields retain their native wheel behavior.
            if isinstance(widget, (tk.Frame, tk.Label, tk.Button, ttk.Frame, ttk.Label)):
                widget.bind("<MouseWheel>", self.wheel)
            pending.extend(widget.winfo_children())


class CompactInputPanel(InputPanel):
    """Display adapter: ordered fields use rows; the existing rules API remains intact."""
    def __init__(self, parent, title, changed, *, fields_parent=None, advanced_parent=None):
        ttk.Frame.__init__(self, parent, padding=px(parent, 8), style="WB.Card.TFrame")
        self.paths, self.controls = [], []
        self.changed = changed
        self.sheet, self.header = tk.StringVar(), tk.StringVar(value="1")
        self.key, self.amount = tk.StringVar(value="客户"), tk.StringVar(value="金额")
        self.extra_keys = tk.StringVar()
        self._extra_variables, self._extra_widgets = [], []
        self._updating = False
        self.disabled = False
        self._input_caption = title
        self.input_title = label(self, title, bold=True)
        self.input_title.pack(anchor="w")
        self.files = ttk.Treeview(self, show="tree", height=2, style="Desktop.FileList.Treeview")
        self.files.column("#0", width=px(self, 190), minwidth=px(self, 100))
        self.files.pack(fill="x", pady=px(self, 5))
        self.controls.append(self.files)
        actions = tk.Frame(self, bg=SURFACE)
        actions.pack(fill="x")
        for caption, action in (("选择文件", self.select_files), ("清空", lambda: self.set_files([]))):
            control = button(actions, caption, action, pady=px(self, 3))
            control.pack(side="left", padx=px(self, (0, 5)))
            self.controls.append(control)
        form = fields_parent if fields_parent is not None else tk.Frame(self, bg=SURFACE)
        if fields_parent is None:
            form.pack(fill="x", pady=px(self, (6, 0)))
        self.field_form = form
        key_heading = tk.Frame(form, bg=SURFACE)
        key_heading.pack(fill="x", pady=px(self, (7, 2)))
        label(key_heading, "键字段①", bold=True).pack(side="left")
        self.add_key_button = button(key_heading, "添加组合字段", self.add_key, pady=0, padx=0,
                                     fg=ACCENT, highlightthickness=0, font=("Microsoft YaHei UI", 9))
        self.add_key_button.pack(side="right")
        self.controls.append(self.add_key_button)
        self.key_entry = ttk.Entry(form, textvariable=self.key, width=17, font=("Microsoft YaHei UI", 12), style="Desktop.TEntry")
        self.key_entry.pack(fill="x")
        self.controls.append(self.key_entry)
        self.extra_host = tk.Frame(form, bg=SURFACE)
        self.extra_host.pack(fill="x")
        label(form, "金额字段", bold=True).pack(anchor="w", pady=px(self, (5, 2)))
        amount_entry = ttk.Entry(form, textvariable=self.amount, width=17, font=("Microsoft YaHei UI", 12), style="Desktop.TEntry")
        amount_entry.pack(fill="x")
        self.controls.append(amount_entry)
        advanced_form = advanced_parent or form
        self.advanced_button = button(advanced_form, "高级设置", self.toggle_advanced, pady=px(self, 3),
                                      highlightthickness=0, fg=SECONDARY)
        self.advanced_button.pack(anchor="w", pady=px(self, 5))
        self.controls.append(self.advanced_button)
        self.advanced = tk.Frame(advanced_form, bg=SURFACE)
        for caption, variable in (("工作表（留空用默认表）", self.sheet), ("表头行（默认 1）", self.header)):
            label(self.advanced, caption, muted=True).pack(anchor="w", pady=px(self, (3, 1)))
            entry = ttk.Entry(self.advanced, textvariable=variable, width=17, font=("Microsoft YaHei UI", 12), style="Desktop.TEntry")
            entry.pack(fill="x")
            self.controls.append(entry)
        label(advanced_form, "精确表头或 #列号，组合最多 8 项。", muted=True,
              wraplength=px(self, 245), size=9).pack(anchor="w", pady=px(self, 5))
        for variable in (self.sheet, self.header, self.key, self.amount):
            variable.trace_add("write", lambda *_: self.changed())
        self.extra_keys.trace_add("write", self._refresh_extra_fields)

    def set_files(self, files):
        self.paths = list(dict.fromkeys(str(Path(p).resolve()) for p in files))
        self.files.delete(*self.files.get_children())
        for path in self.paths:
            self.files.insert("", "end", text=Path(path).name,
                              image=self.winfo_toplevel()._desktop_icons["rules"])
        if hasattr(self, "input_title"):
            self.input_title.configure(text=f"{self._input_caption}  {len(self.paths)} 个文件")
        self.changed()

    def toggle_advanced(self):
        if self.disabled:
            return
        if self.advanced.winfo_manager():
            self.advanced.pack_forget()
        else:
            self.advanced.pack(fill="x")

    def add_key(self):
        if self.disabled:
            return
        data = json.loads(self.extra_keys.get() or "[]")
        if len(data) < 7:
            self.extra_keys.set(json.dumps([*data, ""], ensure_ascii=False))

    def _refresh_extra_fields(self, *_):
        if self._updating:
            return
        try:
            data = json.loads(self.extra_keys.get() or "[]")
            if not isinstance(data, list) or len(data) > 7:
                return  # The unchanged rules() validator reports invalid external input.
        except (ValueError, TypeError):
            return
        for widget in self._extra_widgets:
            if widget in self.controls:
                self.controls.remove(widget)
        for child in self.extra_host.winfo_children():
            child.destroy()
        self._extra_variables, self._extra_widgets = [], []
        for position, text in enumerate(data):
            row = tk.Frame(self.extra_host, bg=SURFACE)
            row.pack(fill="x", pady=px(self, (5, 0)))
            label(row, f"键字段{position + 2}").pack(anchor="w")
            variable = tk.StringVar(value=text)
            self._extra_variables.append(variable)
            entry = ttk.Entry(row, textvariable=variable, width=14, font=("Microsoft YaHei UI", 12), style="Desktop.TEntry")
            entry.pack(side="left", fill="x", expand=True)
            remove = button(row, "移除", lambda index=position: self.remove_key(index), pady=px(self, 1), padx=px(self, 4))
            remove.pack(side="right")
            self.controls.extend((entry, remove))
            self._extra_widgets.extend((entry, remove))
            variable.trace_add("write", self._sync_extra_fields)
        self.changed()

    def _sync_extra_fields(self, *_):
        self._updating = True
        try:
            self.extra_keys.set(json.dumps([item.get() for item in self._extra_variables], ensure_ascii=False))
        finally:
            self._updating = False
        self.changed()

    def remove_key(self, index):
        if not self.disabled:
            data = json.loads(self.extra_keys.get() or "[]")
            del data[index]
            self.extra_keys.set(json.dumps(data, ensure_ascii=False) if data else "")


class EmbeddedWorkbench(WorkbenchController, ttk.Frame):
    def __init__(self, parent, template_path, can_start):
        ttk.Frame.__init__(self, parent, style="WB.TFrame")
        self.can_start = can_start
        self._init_state(template_path, history_writer=task_history.save_task)

    def _build(self):
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        settings = SettingsScroll(self)
        settings.grid(row=0, column=0, sticky="nsew", padx=px(self, (0, 12)))
        settings.configure(width=px(self, 252), highlightthickness=1, highlightbackground="#E4E8F0")
        settings.grid_propagate(False)
        self.settings_scroll = settings
        self.month_config = tk.Frame(settings.inner, bg=SURFACE, padx=px(self, 10), pady=px(self, 8))
        self.compare_config = tk.Frame(settings.inner, bg=SURFACE, padx=px(self, 10), pady=px(self, 8))
        self.month_config.pack(fill="x")
        # A hidden native notebook preserves the controller's established mode API.
        self.tabs = ttk.Notebook(self)
        for caption in ("每月重复汇总", "两份账单核对差异"):
            self.tabs.add(ttk.Frame(self.tabs), text=caption)
        self.template_name = tk.StringVar(value="默认月报")
        self.month = tk.StringVar(value=datetime.date.today().strftime("%Y-%m"))
        self.output_prefix = tk.StringVar(value="月度汇总")
        self.mode, self.tolerance = tk.StringVar(value="客户总额"), tk.StringVar(value="0")
        fields_host = tk.Frame(self.month_config, bg=SURFACE)
        advanced_host = tk.Frame(self.month_config, bg=SURFACE)
        self.monthly_input = CompactInputPanel(self.month_config, "本次输入", self.invalidate,
                                               fields_parent=fields_host, advanced_parent=advanced_host)
        self.monthly_input.pack(fill="x")
        label(self.month_config, "模板", bold=True).pack(anchor="w", pady=px(self, (10, 4)))
        self.template_box = ttk.Combobox(self.month_config, textvariable=self.template_name, width=18,
                                        postcommand=self.refresh_templates, font=("Microsoft YaHei UI", 12), style="Desktop.TCombobox")
        self.template_box.pack(fill="x")
        template_actions = tk.Frame(self.month_config, bg=SURFACE)
        template_actions.pack(fill="x", pady=px(self, 4))
        load_button = button(template_actions, "读取模板", self.load_template, pady=px(self, 2), padx=px(self, 3),
                             highlightthickness=0, fg=ACCENT, font=("Microsoft YaHei UI", 9))
        save_button = button(template_actions, "保存规则", self.save_template, pady=px(self, 2), padx=px(self, 3),
                             highlightthickness=0, fg=ACCENT, font=("Microsoft YaHei UI", 9))
        load_button.pack(side="left")
        save_button.pack(side="left", padx=px(self, 4))
        label(self.month_config, "月份（YYYY-MM）", bold=True).pack(anchor="w", pady=px(self, (7, 4)))
        month_row = tk.Frame(self.month_config, bg=SURFACE)
        month_row.pack(fill="x")
        tk.Label(month_row, image=self.winfo_toplevel()._desktop_icons["calendar"], bg=SURFACE).pack(
            side="left", padx=px(self, (0, 6)))
        month_entry = ttk.Entry(month_row, textvariable=self.month, width=18,
                                font=("Microsoft YaHei UI", 12), style="Desktop.TEntry")
        month_entry.pack(side="left", fill="x", expand=True)
        fields_host.pack(fill="x")
        self.date_filter, self.date_field = tk.BooleanVar(value=False), tk.StringVar()
        self.date_check = ttk.Checkbutton(self.month_config, text="按日期筛选", variable=self.date_filter,
                                          command=self._date_toggle, style="Desktop.TCheckbutton")
        self.date_check.pack(anchor="w", pady=px(self, (8, 2)))
        self.date_entry = ttk.Entry(self.month_config, textvariable=self.date_field, state="disabled", width=18,
                                    font=(self.winfo_toplevel()._display_font_family, 12), style="Desktop.TEntry")
        label(self.month_config, "关闭时，月份仅作标签。\n开启时请填写日期字段。", muted=True,
              wraplength=px(self, 245)).pack(anchor="w", pady=px(self, 4))
        output_options = tk.Frame(self.month_config, bg=SURFACE)
        def toggle_output():
            if output_options.winfo_manager():
                output_options.pack_forget()
            else:
                output_options.pack(fill="x")
        output_button = button(self.month_config, "输出设置", toggle_output, pady=px(self, 3))
        output_button.pack(anchor="w", pady=px(self, 4))
        advanced_host.pack(fill="x")
        label(output_options, "输出文件前缀").pack(anchor="w", pady=px(self, (6, 2)))
        prefix_entry = ttk.Entry(output_options, textvariable=self.output_prefix, width=18, font=("Microsoft YaHei UI", 12), style="Desktop.TEntry")
        prefix_entry.pack(fill="x")
        label(self.compare_config, "对账模式").pack(anchor="w", pady=px(self, (0, 3)))
        mode_box = ttk.Combobox(self.compare_config, textvariable=self.mode,
                               values=("客户总额", "单据逐笔"), state="readonly", width=18, font=("Microsoft YaHei UI", 12), style="Desktop.TCombobox")
        mode_box.pack(fill="x")
        mode_box.bind("<<ComboboxSelected>>", self.mode_changed)
        label(self.compare_config, "金额容差（默认 0）").pack(anchor="w", pady=px(self, (6, 2)))
        tolerance_entry = ttk.Entry(self.compare_config, textvariable=self.tolerance, width=18, font=("Microsoft YaHei UI", 12), style="Desktop.TEntry")
        tolerance_entry.pack(fill="x")
        self.left_input = CompactInputPanel(self.compare_config, "左侧账单", self.invalidate)
        self.right_input = CompactInputPanel(self.compare_config, "右侧账单", self.invalidate)
        self.left_input.pack(fill="x", pady=px(self, (8, 3)))
        self.right_input.pack(fill="x", pady=px(self, 3))
        label(self.compare_config, "组合顺序一一对应；重复单据列歧义，不自动配对。", muted=True,
              wraplength=px(self, 245)).pack(anchor="w", pady=px(self, 7))
        self.controls = [(widget, str(widget.cget("state"))) for widget in
                         (self.template_box, load_button, save_button, month_entry, prefix_entry,
                          self.date_check, mode_box, tolerance_entry, output_button)]
        right = tk.Frame(self, bg=SURFACE, padx=px(self, 16), pady=px(self, 14),
                         highlightthickness=1, highlightbackground="#E4E8F0")
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(2, weight=1)
        right.columnconfigure(0, weight=1)
        label(right, "结果预览", size=16, bold=True).grid(row=0, column=0, sticky="w", pady=px(self, (0, 12)))
        self.empty_hint = label(right, "请选择本次文件与规则，再预览。不会自动加载示例或旧账单。", muted=True,
                                wraplength=px(self, 650))
        self.empty_hint.grid(row=1, column=0, sticky="w", pady=px(self, (0, 8)))
        self.empty_hint.configure(bg="#EFF6FF", fg=ACCENT, padx=px(self, 10), pady=px(self, 8), anchor="w")
        self.empty_hint.grid_configure(sticky="ew")
        self.views = ttk.Notebook(right, style="Desktop.TNotebook")
        self.views.grid(row=2, column=0, sticky="nsew")
        self.view_indicator = tk.Frame(self.views, bg=ACCENT, height=px(self, 3))
        self.views.bind("<Configure>", self._place_view_indicator, add="+")
        self.views.bind("<<NotebookTabChanged>>", self._place_view_indicator, add="+")
        self.summary = DesktopTree(self.views, ("键", "分类", "汇总 / 左金额", "右金额", "左减右", "来源行数"))
        for index, width in enumerate((150, 90, 130, 130, 130, 85)):
            self.summary.tree.column(f"c{index}", width=px(self, width), minwidth=px(self, 75))
        self.summary.tree.configure(displaycolumns=("c0", "c2", "c5"))
        self.summary.tree.heading("c2", text="金额（精确文本）")
        self.sources = DesktopTree(self.views, ("侧别", "分类", "文件", "工作表", "原行号", "键", "精确金额", "日期（原｜规范）"))
        self.issues = DesktopTree(self.views, ("侧别", "文件", "工作表", "原行号", "异常", "说明", "原始键", "原始金额", "原日期", "规范日期"))
        for frame, caption in ((self.summary, "汇总结果"), (self.sources, "来源明细"), (self.issues, "异常清单")):
            self.views.add(frame, text=caption)
        self.status = tk.StringVar(value="尚未运行；全部处理在本机完成。")
        self.status_label = label(right, "", muted=True, wraplength=px(self, 650))
        self.status_label.configure(textvariable=self.status)
        self.status_label.grid(row=3, column=0, sticky="ew", pady=px(self, (8, 0)))
        right.bind("<Configure>", lambda event: self.status_label.configure(
            wraplength=max(px(self, 140), event.width - px(self, 30))), add="+")
        footer = tk.Frame(self, bg=BG, pady=px(self, 8))
        footer.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.preview_button = button(footer, "预览汇总", self.preview)
        self.preview_button.pack(side="right", padx=px(self, (0, 10)))
        self.cancel_button = button(footer, "取消", self.cancel, state="disabled")
        self.cancel_button.pack(side="left", padx=px(self, 6))
        self.export_button = button(footer, "导出结果", self.export, primary=True, state="disabled")
        self.export_button.pack(side="right")
        self.preview_button.pack_forget()
        self.preview_button.pack(side="right", padx=px(self, (0, 10)))
        self._active_tab = self.tabs.select()
        self.tabs.bind("<<NotebookTabChanged>>", self.tab_changed)
        for variable in (self.template_name, self.month, self.output_prefix, self.mode, self.tolerance, self.date_field):
            variable.trace_add("write", lambda *_: self.invalidate())
        settings.bind_form()

    def select_mode(self, mode):
        if self._busy or self.runner.busy:
            return False
        compare = mode == "compare"
        self.tabs.select(1 if compare else 0)
        self.month_config.pack_forget()
        self.compare_config.pack_forget()
        (self.compare_config if compare else self.month_config).pack(fill="x")
        self.preview_button.configure(text="预览对账" if compare else "预览汇总")
        self.summary.tree.configure(displaycolumns=tuple(f"c{index}" for index in range(6))
                                    if compare else ("c0", "c2", "c5"))
        # Tk stretches visible columns; reset before changing between 3 and 6 columns.
        widths = tuple(px(self, width) for width in (150, 90, 130, 130, 130, 85))
        visible = range(6) if compare else (0, 2, 5)
        scale = max(1, (self.summary.tree.winfo_width() - px(self, 22)) / sum(widths[index] for index in visible))
        for index, width in enumerate(widths):
            self.summary.tree.column(f"c{index}", width=int(width * scale) if index in visible else width, minwidth=px(self, 75))
        self.summary.tree.heading("c2", text="左金额" if compare else "金额（精确文本）")
        self.settings_scroll.offset = 0
        self.settings_scroll.layout()
        self.tab_changed()
        return True

    def _place_view_indicator(self, _event=None):
        """A normal tab underline; the notebook retains navigation and focus."""
        if not self.views.tabs() or not self.views.select():
            return
        selected = self.views.index(self.views.select())
        def index_at(x, y):
            try:
                return self.views.index(f"@{x},{y}")
            except tk.TclError:
                return None
        # Query the real themed tab boundaries, not estimated font widths.
        y = px(self, 12)
        bounds = [x for x in range(self.views.winfo_width()) if index_at(x, y) == selected]
        if not bounds:
            self.view_indicator.place_forget()
            return
        left, right = bounds[0], bounds[-1] + 1
        center = (left + right) // 2
        bottom = y
        while bottom < px(self, 80) and index_at(center, bottom) == selected:
            bottom += 1
        thickness = px(self, 3)
        self.view_indicator.place(x=left, y=bottom - thickness, width=right - left, height=thickness)
        self.view_indicator.lift()

    def _set_busy(self, value):
        super()._set_busy(value)
        for panel in (self.monthly_input, self.left_input, self.right_input):
            panel.disabled = value
            for widget in panel.controls:
                if isinstance(widget, ttk.Treeview):
                    widget.state(("disabled",) if value else ("!disabled",))
                else:
                    widget.configure(state="disabled" if value else "normal")

    def preview(self):
        if self.can_start():
            super().preview()

    def invalidate(self, *_):
        super().invalidate()
        if hasattr(self, "empty_hint"):
            self.empty_hint.configure(text="请选择本次文件与规则，再预览。旧结果不会用于导出。")

    def _date_toggle(self):
        super()._date_toggle()
        if self.date_filter.get():
            self.date_entry.pack(after=self.date_check, fill="x", pady=px(self, 3))
        else:
            self.date_entry.pack_forget()

    def export(self):
        if self.can_start():
            super().export()

    def _show_result(self, result):
        super()._show_result(result)
        from .monthly import MonthlyResult
        text = (f"金额合计（精确）：{result.total}。" if isinstance(result, MonthlyResult) else "")
        self.empty_hint.configure(text=text + "请核对来源与异常；分页不影响完整导出。")


class EmbeddedSummary(gui.SummaryApp):
    """Only builds an alternative legacy view; legacy business methods are inherited."""
    def __init__(self, root, host, can_start):
        self.can_start = can_start
        super().__init__(root, host=host)
        self._show_task()

    def _build_welcome(self):
        inner = self._create_scroll_host()
        self.welcome_frame = ttk.Frame(inner, style="Page.TFrame")

    def _build_task_frame(self):
        self.task_frame = ttk.Frame(self._scroll_inner, padding=px(self, 8), style="Page.TFrame")
        self.task_frame.columnconfigure(1, weight=1)
        self.task_frame.rowconfigure(0, weight=1)
        left = tk.Frame(self.task_frame, bg=SURFACE, padx=px(self, 12), pady=px(self, 12), width=px(self, 275))
        left.grid(row=0, column=0, sticky="nsew", padx=px(self, (0, 12)))
        right = tk.Frame(self.task_frame, bg=SURFACE, padx=px(self, 12), pady=px(self, 12))
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        self._responsive_hints, self._nav_hint, self._action_hint = [], None, None
        self._responsive_pending = False
        label(left, "输入文件夹").pack(anchor="w")
        self.folder_entry = ttk.Entry(left, textvariable=self.selected_folder, width=22,
                                      font=(self.root._display_font_family, 12), style="Desktop.TEntry")
        self.folder_entry.pack(fill="x", pady=px(self, 5))
        self.browse_button = button(left, "浏览文件夹", self._on_browse_folder)
        self.browse_button.pack(anchor="w")
        label(left, "保留原自动识别流程，\n只处理本次勾选的文件。", muted=True).pack(anchor="w", pady=px(self, 10))
        self.rules_button = button(left, "管理字段规则", self._on_open_field_config)
        self.rules_button.pack(anchor="w")
        label(left, "输出位置").pack(anchor="w", pady=px(self, (18, 3)))
        self._output_row = tk.Frame(left, bg=SURFACE)
        self._output_row.pack(fill="x")
        self.output_radios = []
        for caption, value in (("默认输出目录", "default"), ("输入文件夹", "input"), ("自定义目录", "custom")):
            radio = tk.Radiobutton(self._output_row, text=caption, variable=self.output_mode, value=value,
                                   bg=SURFACE, fg=TEXT, activebackground=SURFACE, font=("Microsoft YaHei UI", 10))
            radio.pack(anchor="w")
            self.output_radios.append(radio)
        self._custom_dir_btn = button(self._output_row, "选择目录", self._on_select_custom_dir)
        self.output_mode.trace_add("write", self._on_output_mode_changed)
        self._sync_custom_dir_button()
        self.start_btn = button(left, "开始汇总", self._on_start_summary, primary=True, state="disabled")
        self.start_btn.pack(anchor="w", pady=px(self, (18, 5)))
        label(right, "文件检查与运行结果", title=True).grid(row=0, column=0, sticky="w", pady=px(self, (0, 8)))
        self.file_status_label = label(right, "请先选择文件夹", muted=True, wraplength=px(self, 500))
        self.file_status_label.grid(row=1, column=0, sticky="ew")
        list_row = tk.Frame(right, bg=SURFACE)
        list_row.grid(row=2, column=0, sticky="ew", pady=px(self, 8))
        list_row.columnconfigure(0, weight=1)
        self.file_list_text = gui._ui_text_field(list_row, height=6, width=30, state="disabled")
        self.file_list_text.grid(row=0, column=0, sticky="ew")
        self._file_list_scrollbar = ttk.Scrollbar(list_row, orient="vertical", command=self.file_list_text.yview)
        self._file_list_scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_list_text.configure(yscrollcommand=self._on_file_list_scrolled)
        self.file_list_text.bind("<MouseWheel>", self._on_file_list_wheel)
        self._file_list_line_count = 0
        self._status_border = tk.Frame(right, bg=SURFACE)
        self._status_border.grid(row=3, column=0, sticky="ew", pady=px(self, 8))
        self._status_card = self._status_border
        self._status_card.columnconfigure(0, weight=1)
        label(self._status_card, "运行详情").grid(row=0, column=0, sticky="w")
        self.status_text = gui._ui_text_field(self._status_card, height=5, width=30, state="disabled", wrap="word")
        self.status_text.grid(row=1, column=0, sticky="ew", pady=px(self, 4))
        self.progress = ttk.Progressbar(self._status_card, mode="indeterminate")
        result_row = tk.Frame(right, bg=SURFACE)
        result_row.grid(row=4, column=0, sticky="ew", pady=px(self, 8))
        self.open_btn = button(result_row, "打开汇总文件", self._on_open_file, state="disabled")
        self.open_folder_btn = button(result_row, "打开文件位置", self._on_open_folder, state="disabled")
        self.open_btn.pack(side="left")
        self.open_folder_btn.pack(side="left", padx=px(self, 6))
        self._bind_page_wheel(self.task_frame)

    def _on_start_summary(self):
        if self.can_start():
            super()._on_start_summary()


class UnifiedApp:
    """A single desktop window and a single business writer, including across pages."""
    MODES = {"regular": "常规汇总", "monthly": "月度汇总", "compare": "账单对账"}

    def __init__(self, root):
        if not hasattr(root, "_display_density"):
            configure_display(root)
        configure_styles(root)
        self.root, self._closing, self._after = root, False, None
        self._destroyed_root = False
        self.workbench, self.legacy = None, None
        self.active_page, self.active_mode = "workbench", "monthly"
        root.title(APP_FULL_NAME)
        apply_app_icon(root)
        root.configure(bg=BG)
        initial, minimum = window_bounds(root)
        root.geometry(f"{initial[0]}x{initial[1]}")
        root.minsize(*minimum)
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.bind("<Destroy>", self._destroyed, add="+")
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        icon_size = min(ICON_PNGS, key=lambda size: abs(size - px(root, 24)))
        brand_size = min(BRAND_PNGS, key=lambda size: abs(size - px(root, 32)))
        self.icons = {name: tk.PhotoImage(master=root, data=data) for name, data in ICON_PNGS[icon_size]["normal"].items()}
        self.active_icons = {name: tk.PhotoImage(master=root, data=data) for name, data in ICON_PNGS[icon_size]["active"].items()}
        self.icons["brand"] = tk.PhotoImage(master=root, data=BRAND_PNGS[brand_size])
        root._desktop_icons = self.icons
        nav = tk.Frame(root, bg=SURFACE, width=px(self, 190), padx=px(self, 10), pady=px(self, 16))
        nav.grid(row=0, column=0, sticky="ns")
        nav.grid_propagate(False)
        nav.columnconfigure(0, weight=1)
        tk.Label(nav, text="  " + APP_NAME, image=self.icons["brand"], compound="left", bg=SURFACE,
                 fg=TEXT, font=("Microsoft YaHei UI", 17, "bold")).grid(
                     row=0, column=0, sticky="w", padx=px(self, 10), pady=px(self, (6, 22)))
        self.nav_buttons = {}
        self.nav_indicators = {}
        for row, (key, caption) in enumerate((("workbench", "工作台"), ("history", "任务记录"),
                                              ("rules", "规则与模板"), ("help", "使用教程"),
                                              ("feedback", "问题反馈")), 1):
            slot = tk.Frame(nav, bg=SURFACE)
            slot.grid(row=row, column=0, sticky="ew", pady=px(self, 3))
            mark = tk.Frame(slot, bg=SURFACE, width=px(self, 3))
            mark.pack(side="left", fill="y")
            control = button(slot, "  " + caption, lambda name=key: self.show_page(name), anchor="w",
                             image=self.icons[key], compound="left", highlightthickness=0, pady=px(self, 11))
            control.pack(side="left", fill="x", expand=True)
            self.nav_buttons[key] = control
            self.nav_indicators[key] = mark
        nav.rowconfigure(6, weight=1)
        button(nav, "  关于 E助手", lambda: self.show_page("about"), anchor="w", image=self.icons["about"],
               compound="left", highlightthickness=0).grid(row=7, column=0, sticky="ew")
        tk.Label(nav, text="  离线运行", image=self.icons["offline"], compound="left", bg=SURFACE,
                 fg=SECONDARY, font=("Microsoft YaHei UI", 12), anchor="w").grid(row=8, column=0, sticky="w", padx=px(self, 8), pady=px(self, (12, 5)))
        self.main = tk.Frame(root, bg=BG, padx=px(self, 16), pady=px(self, 12))
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.rowconfigure(2, weight=1)
        self.main.columnconfigure(0, weight=1)
        heading = tk.Frame(self.main, bg=BG)
        heading.grid(row=0, column=0, sticky="ew", pady=px(self, (0, 8)))
        self.title = label(heading, "月度汇总", title=True)
        self.title.configure(bg=BG)
        self.title.pack(side="left")
        self.help_button = button(heading, "查看本功能教程  ", self.show_context_help, fg=ACCENT,
                                  image=self.active_icons["external"], compound="right", bg=BG, highlightthickness=0)
        self.help_button.pack(side="right")
        self.mode_bar = tk.Frame(self.main, bg=BG)
        self.mode_bar.grid(row=1, column=0, sticky="ew", pady=px(self, (0, 10)))
        self.mode_buttons = {}
        self.mode_indicators = {}
        for key, caption in self.MODES.items():
            slot = tk.Frame(self.mode_bar, bg=BG)
            slot.pack(side="left", padx=px(self, (0, 12)))
            control = button(slot, caption, lambda name=key: self.select_mode(name), bg=BG, highlightthickness=0)
            control.pack()
            mark = tk.Frame(slot, bg=BG, height=px(self, 3))
            mark.pack(fill="x")
            self.mode_buttons[key] = control
            self.mode_indicators[key] = mark
        self.page_host = tk.Frame(self.main, bg=BG)
        self.page_host.grid(row=2, column=0, sticky="nsew")
        self.pages = {}
        workbench_page = tk.Frame(self.page_host, bg=BG)
        self.pages["workbench"] = workbench_page
        self.regular_host = tk.Frame(workbench_page, bg=BG)
        self.workbench = EmbeddedWorkbench(workbench_page, Path(gui.get_config_dir()) / "monthly_templates.json",
                                           lambda: not self.legacy_busy() and not self._closing)
        self.workbench.pack(fill="both", expand=True)
        self.legacy = EmbeddedSummary(root, self.regular_host,
                                      lambda: not self.workbench_busy() and not self._closing)
        configure_styles(root)
        self.show_page("workbench")
        self.select_mode("monthly")
        self._sync_controls()

    def legacy_busy(self):
        return bool(self.legacy and (self.legacy.task_running or
                    (self.legacy.worker is not None and self.legacy.worker.is_alive())))

    def workbench_busy(self):
        return bool(self.workbench and (self.workbench._busy or self.workbench.runner.busy))

    def busy(self):
        return self._closing or self.legacy_busy() or self.workbench_busy()

    def select_mode(self, mode):
        if mode not in self.MODES or self.busy():
            return False
        self.active_mode = mode
        self.workbench.pack_forget()
        self.regular_host.pack_forget()
        if mode == "regular":
            self.regular_host.pack(fill="both", expand=True)
        else:
            self.workbench.select_mode(mode)
            self.workbench.pack(fill="both", expand=True)
        self.show_page("workbench")
        self.title.configure(text=self.MODES[mode])
        for key, control in self.mode_buttons.items():
            control.configure(bg=BG, fg=ACCENT if key == mode else TEXT,
                              font=("Microsoft YaHei UI", 12, "bold" if key == mode else "normal"))
            self.mode_indicators[key].configure(bg=ACCENT if key == mode else BG)
        return True

    def show_page(self, page):
        if page not in ("workbench", "history", "rules", "help", "feedback", "about"):
            return False
        if self.busy() and page in ("history", "rules"):
            return False
        if page not in self.pages:
            self.pages[page] = getattr(self, "_build_" + page)()
        for frame in self.pages.values():
            frame.pack_forget()
        self.pages[page].pack(fill="both", expand=True)
        self.active_page = page
        if page == "workbench":
            self.mode_bar.grid()
            self.help_button.pack(side="right")
            self.title.configure(text=self.MODES[self.active_mode])
        else:
            self.mode_bar.grid_remove()
            self.help_button.pack_forget()
            self.title.configure(text={"history": "任务记录", "rules": "规则与模板", "help": "使用教程",
                                       "feedback": "问题反馈", "about": "关于 E助手"}[page])
        for key, control in self.nav_buttons.items():
            control.configure(bg="#E8F1FE" if key == page else SURFACE, fg=ACCENT if key == page else TEXT,
                              image=self.active_icons[key] if key == page else self.icons[key],
                              font=(self.root._display_font_family, 12, "bold" if key == page else "normal"))
            self.nav_indicators[key].configure(bg=ACCENT if key == page else SURFACE)
        if page == "history":
            self.refresh_history()
        return True

    def _build_history(self):
        frame = tk.Frame(self.page_host, bg=SURFACE, padx=px(self, 16), pady=px(self, 14))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(3, weight=1)
        label(frame, "本地任务记录", title=True).grid(row=0, column=0, sticky="w")
        label(frame, "常规汇总、月度汇总与两种对账；预览不等于已导出。仅存本机元数据，最多500条。", muted=True,
              wraplength=px(self, 650)).grid(row=1, column=0, sticky="w", pady=px(self, 8))
        filters = tk.Frame(frame, bg=SURFACE)
        filters.grid(row=2, column=0, sticky="ew", pady=px(self, (0, 8)))
        label(filters, "任务类型").pack(side="left", padx=px(self, (0, 8)))
        self.history_filter = tk.StringVar(value="全部")
        self.history_filter_box = ttk.Combobox(filters, textvariable=self.history_filter, state="readonly",
                                              values=("全部", *task_history.TASK_TYPES.values()), width=18,
                                              font=(self.root._display_font_family, 12), style="Desktop.TCombobox")
        self.history_filter_box.pack(side="left")
        self.history_filter_box.bind("<<ComboboxSelected>>", lambda _event: self.refresh_history())
        self.history_tree = HistoryTree(frame, ("时间", "任务类型", "状态", "文件数", "结果数量", "输出文件"))
        self.history_tree.grid(row=3, column=0, sticky="nsew")
        self.history_records = []
        self.history_tree.tree.bind("<<TreeviewSelect>>", self._history_selection)
        self.history_notice = label(frame, "", muted=True, wraplength=px(self, 650))
        self.history_notice.grid(row=4, column=0, sticky="ew", pady=px(self, 6))
        # Parent width includes horizontal padding; keep wrapped text inside the grid.
        frame.bind("<Configure>", lambda event: self.history_notice.configure(
            wraplength=max(px(self, 140), min(px(self, 650), event.width - px(self, 40)))), add="+")
        actions = tk.Frame(frame, bg=SURFACE)
        actions.grid(row=5, column=0, sticky="ew", pady=px(self, 8))
        button(actions, "刷新", self.refresh_history).pack(side="left")
        self.history_open_button = button(actions, "打开结果", self.open_history_result, state="disabled")
        self.history_open_button.pack(side="left", padx=px(self, 8))
        button(actions, "管理记录", self.open_history).pack(side="left")
        return frame

    def refresh_history(self):
        if not self.busy():
            records, warning = task_history.read_history()
            kind = next((key for key, text in task_history.TASK_TYPES.items()
                         if text == self.history_filter.get()), None)
            self.history_records = list(reversed(task_history.filter_records(records, kind)))
            self.history_tree.set_rows((record.get("time", ""), task_history.type_label(record),
                                       task_history.status_label(record), record.get("file_count", 0),
                                       record.get("result_count", "未记录"), record.get("output_file", ""))
                                      for record in self.history_records)
            self.history_notice.configure(text=warning or f"当前{len(self.history_records)}条记录。选中可查看输出位置；文件已移动或删除时不能直接打开。")
            self._history_selection()

    def _selected_history(self):
        selected = self.history_tree.tree.selection()
        if not selected:
            return None
        try:
            index = self.history_tree.page * self.history_tree.PAGE_SIZE + self.history_tree.tree.index(selected[0])
            return self.history_records[index] if index < len(self.history_records) else None
        except tk.TclError:
            return None

    def _history_selection(self, _event=None):
        record = self._selected_history()
        path = task_history.output_path(record) if record else ""
        self.history_open_button.configure(state="normal" if path and not self.busy() else "disabled")
        if record:
            self.history_notice.configure(text=f"输出位置：{path}" if path else "本条记录没有已生成的Excel；预览不等于已导出。")

    def open_history_result(self):
        if self.busy():
            return
        record = self._selected_history()
        path = task_history.output_path(record) if record else ""
        if not path or not os.path.isfile(path):
            messagebox.showinfo("无法打开结果", "本条记录没有已生成的本地Excel，或文件已被移动/删除。", parent=self.root)
            return
        try:
            os.startfile(path)
        except OSError:
            messagebox.showwarning("无法打开结果", "系统未能打开文件，请检查Excel关联或手动打开输出位置。", parent=self.root)

    def open_history(self):
        if not self.busy():
            gui.HistoryDialog(self.root)

    def _build_rules(self):
        frame = tk.Frame(self.page_host, bg=SURFACE, padx=px(self, 24), pady=px(self, 20))
        label(frame, "规则分开保存，业务口径不混用", title=True).pack(anchor="w", pady=px(self, 10))
        label(frame, "常规汇总：原自动识别字段与人工规则。\n月度汇总：命名规则模板，每次重新选文件。\n"
                     "账单对账：两侧字段独立设置，按顺序精确对应。", wraplength=px(self, 650)).pack(anchor="w", pady=px(self, 12))
        button(frame, "管理常规字段规则", self.open_field_rules).pack(anchor="w", pady=px(self, 6))
        button(frame, "前往月度模板", lambda: self.select_mode("monthly")).pack(anchor="w", pady=px(self, 6))
        button(frame, "设置对账字段", lambda: self.select_mode("compare")).pack(anchor="w", pady=px(self, 6))
        return frame

    def open_field_rules(self):
        if not self.busy():
            self.legacy._on_open_field_config()

    def _build_help(self):
        frame = tk.Frame(self.page_host, bg=SURFACE, padx=px(self, 12), pady=px(self, 12))
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(1, weight=1)
        label(frame, "内置教程 · 离线可用", muted=True).grid(row=0, column=0, columnspan=2, sticky="w", pady=px(self, (0, 8)))
        topics = tk.Frame(frame, bg=SURFACE)
        topics.grid(row=1, column=0, sticky="ns", padx=px(self, (0, 12)))
        self.help_query = tk.StringVar()
        ttk.Entry(topics, textvariable=self.help_query, width=22, font=(self.root._display_font_family, 12),
                  style="Desktop.TEntry").pack(fill="x", pady=px(self, (0, 8)))
        self.help_list = tk.Listbox(topics, width=22, height=12, font=("Microsoft YaHei UI", 10),
                                   relief="flat", exportselection=False, bg="#FAFAFB")
        self.help_list.pack(fill="both", expand=True)
        self.help_list.bind("<<ListboxSelect>>", self._select_article)
        body = tk.Frame(frame, bg=SURFACE)
        body.grid(row=1, column=1, sticky="nsew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.help_text = tk.Text(body, wrap="word", relief="flat", bg=SURFACE, fg=TEXT,
                                 font=("Microsoft YaHei UI", 12), padx=px(self, 12), pady=px(self, 8), spacing1=px(self, 4), spacing3=px(self, 6),
                                 width=36, state="disabled")
        self.help_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(body, command=self.help_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.help_text.configure(yscrollcommand=scrollbar.set)
        self.article_button = button(body, "前往相关功能", self.open_article_mode)
        self.article_button.grid(row=1, column=0, sticky="w", pady=px(self, 8))
        self.help_query.trace_add("write", self._search_help)
        self._search_help()
        return frame

    def _search_help(self, *_):
        self.visible_articles = search_articles(self.help_query.get())
        self.help_list.delete(0, "end")
        for article in self.visible_articles:
            self.help_list.insert("end", article.title)
        self.selected_article = self.visible_articles[0] if self.visible_articles else None
        if self.visible_articles:
            self.help_list.selection_set(0)
        self._render_article()

    def _select_article(self, _event=None):
        selected = self.help_list.curselection()
        if selected:
            self.selected_article = self.visible_articles[selected[0]]
            self._render_article()

    def _render_article(self):
        article = self.selected_article
        self.help_text.configure(state="normal")
        self.help_text.delete("1.0", "end")
        self.help_text.insert("end", article.title + "\n\n" + article.body if article else "没有匹配的内置教程，请换一个关键词。")
        self.help_text.configure(state="disabled")
        self.article_button.configure(state="normal" if article and not self.busy() else "disabled")

    def open_article_mode(self):
        if self.selected_article and not self.busy():
            self.select_mode(self.selected_article.mode)

    def show_context_help(self):
        self.show_page("help")
        title = {"regular": "常规汇总", "monthly": "月度汇总", "compare": "客户总额对账"}[self.active_mode]
        self.help_query.set(title)

    def _build_feedback(self):
        frame = tk.Frame(self.page_host, bg=SURFACE, padx=px(self, 28), pady=px(self, 24))
        label(frame, "联系与反馈", title=True).pack(anchor="w", pady=px(self, (8, 14)))
        label(frame, "软件没有线上服务，不会提交、上传或自动发送任何内容。", muted=True,
              wraplength=px(self, 700)).pack(anchor="w", pady=px(self, 8))
        self.contact_label = label(frame, "\n".join(f"{kind}：{value}" for kind, value in PUBLIC_CONTACTS)
                                   if PUBLIC_CONTACTS else "联系方式暂未提供", wraplength=px(self, 700))
        self.contact_label.pack(anchor="w", pady=px(self, 24))
        label(frame, "后续由维护者提供公开联系方式。\n分享问题前，请遮挡账单、个人信息、文件路径、授权信息或密钥。\n"
                     "不要发送API配置或真实业务文件。", muted=True, wraplength=px(self, 700)).pack(anchor="w", pady=px(self, 12))
        return frame

    def _build_about(self):
        frame = tk.Frame(self.page_host, bg=SURFACE, padx=px(self, 28), pady=px(self, 24))
        label(frame, APP_NAME, title=True).pack(anchor="w", pady=px(self, 12))
        label(frame, APP_VERSION).pack(anchor="w")
        label(frame, "全部功能离线运行\n无登录、联网检查、在线客服、遥测或自动上传。\n"
                     "输入文件只读；结果与规则保存在本机，目录依赖保持原位。", wraplength=px(self, 700)).pack(anchor="w", pady=px(self, 18))
        return frame

    def _sync_controls(self):
        busy = self.busy()
        for control in self.mode_buttons.values():
            control.configure(state="disabled" if busy else "normal")
        for key in ("history", "rules"):
            self.nav_buttons[key].configure(state="disabled" if busy else "normal")
        if "help" in self.pages:
            self.article_button.configure(state="normal" if self.selected_article and not busy else "disabled")
        if self.legacy is not None:
            for control in (self.legacy.browse_button, self.legacy.rules_button,
                            self.legacy.folder_entry, *self.legacy.output_radios, self.legacy._custom_dir_btn):
                control.configure(state="disabled" if busy else "normal")
        self._after = self.root.after(100, self._sync_controls)

    def close(self):
        if self._destroyed_root or self._closing:
            return
        if self.legacy_busy():
            messagebox.showinfo("任务仍在运行", "常规汇总暂不支持取消，请等待当前任务完成后关闭。", parent=self.root)
            return
        if self.workbench_busy():
            if not messagebox.askyesno("安全关闭", "取消当前任务并等待资源释放后关闭？", parent=self.root):
                return
            self._closing = True
            self.workbench.cancel()
            self._wait_close()
        else:
            self.root.destroy()

    def _wait_close(self):
        if self._destroyed_root:
            return
        if self.workbench.runner.busy:
            self.root.after(30, self._wait_close)
        else:
            if not self.workbench.flush_history():
                messagebox.showwarning("任务记录未保存", self.workbench._history_warning, parent=self.root)
            self.root.destroy()

    def _destroyed(self, event):
        if event.widget is self.root:
            self._destroyed_root = True
            self._closing = True
            if self._after is not None:
                try:
                    self.root.after_cancel(self._after)
                except tk.TclError:
                    pass
                self._after = None
