"""
Excel智能汇总助手 v1.5.4 — GUI 界面

启动方式：
    python -m Excel自动汇总助手.gui
"""

import os
import io
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import font as tkfont
from contextlib import redirect_stdout

# ---- 核心模块导入（只调不修改） ----
from .utils import (
    get_base_dir, get_data_dir, get_config_dir,
    get_logs_dir, get_unique_file_path,
)
from .config import APP_NAME, APP_VERSION, APP_FULL_NAME, FOLDER_NAME, OUTPUT_FILE, OUTPUT_FOLDER_NAME, ERROR_LOG
from .excel_reader import read_and_aggregate, aggregate_with_report, _detect_columns_xls, _get_excel_files
from .field_mapper import (
    detect_columns, get_header_list, get_file_headers,
    save_field_mapping, diagnose_failure, load_settings,
    load_field_config, get_field_config, save_field_config,
    reset_field_config,
)
from .processor import sort_by_amount
from .excel_writer import generate_summary

# ---- v1.5.2: 任务记录 ----
from .history import add_record, load_history, delete_record, clear_records
from . import history as task_history

# ============================================================
#  后台工作线程
# ============================================================


class _WorkerThread(threading.Thread):
    """在后台执行：读取 → 排序 → 写文件，通过队列报告进度。"""

    def __init__(self, folder, output_path, base_dir, file_names=None, field_mappings=None):
        super().__init__(daemon=True)
        self.folder = folder
        self.output_path = output_path
        self.base_dir = base_dir
        # v1.5.2 Hotfix3: 只汇总用户勾选的文件；None 表示全部
        self.file_names = file_names
        # v1.5.4: 当前任务人工字段映射 {文件名: {表头: 字段名}}
        self.field_mappings = field_mappings or {}
        self.queue = queue.Queue()

    def run(self):
        captured = io.StringIO()
        try:
            with redirect_stdout(captured):
                self._pipeline()
        except Exception as exc:
            self.queue.put(("error", str(exc)))
        finally:
            # 把捕获的 print 输出逐行发回 GUI
            for line in captured.getvalue().splitlines():
                if line.strip():
                    self.queue.put(("log", line))

    def _pipeline(self):
        self.queue.put(("status", "正在读取文件..."))

        report = aggregate_with_report(
            self.folder, self.base_dir, self.file_names, self.field_mappings,
        )
        total = report["total"]
        file_count = report["file_count"]
        field_names = report["field_names"]
        skipped = report["skipped"]

        if file_count == 0:
            self.queue.put(("error", "没有可汇总的 Excel 文件"))
            return

        # v1.5.4 核心字段完整性保护：任一选中文件未成功处理 → 阻止生成汇总
        if skipped:
            self.queue.put(("error", _build_skip_error(report)))
            return

        if len(total) == 0:
            self.queue.put(("error", "所有文件均无法读取，没有可汇总的数据"))
            return

        group_name = field_names[0] if len(field_names) >= 1 else "客户"
        self.queue.put(("status", f"已读取 {file_count} 个文件，共 {len(total)} 个{group_name}"))

        self.queue.put(("status", "正在汇总..."))
        sorted_data = sort_by_amount(total, field_names)

        self.queue.put(("status", "正在生成结果文件..."))
        try:
            generate_summary(sorted_data, total, self.output_path, field_names)
        except PermissionError:
            self.queue.put(("file_locked", self.output_path))
            return
        except Exception as e:
            import traceback
            error_log_path = os.path.join(self.base_dir, ERROR_LOG)
            with open(error_log_path, "a", encoding="utf-8") as logf:
                logf.write(f"保存失败 {self.output_path}：{repr(e)}\n{traceback.format_exc()}\n")
            self.queue.put(("error", f"保存失败：{e}"))
            return

        # 计算求和字段的合计
        sum_field = field_names[1] if len(field_names) >= 2 else "金额"
        sum_total = sum(
            v.get(sum_field, 0) if isinstance(v, dict) else v
            for v in total.values()
            if isinstance(v.get(sum_field, 0) if isinstance(v, dict) else v, (int, float))
        )
        self.queue.put(("complete", {
            "file_count": file_count,
            "customer_count": len(total),
            "sum_money": sum_total,
            "output_path": self.output_path,
            "field_names": field_names,
        }))


def _build_skip_error(report):
    """构建核心字段完整性保护错误信息（v1.5.4）。

    当 aggregate_with_report 返回的 skipped 非空时调用，向用户说明哪些
    文件未成功处理以及原因，提示返回字段设置后重新汇总。
    """
    skipped = report["skipped"]
    file_count = report["file_count"]
    processed = report.get("processed_count", 0)

    lines = [
        "汇总无法继续，检测到部分文件未成功处理。",
        f"已选择：{file_count} 个文件",
        f"成功处理：{processed} 个文件",
        f"未处理：{len(skipped)} 个文件",
        "",
        "未处理文件：",
    ]
    for item in skipped:
        fname = item["file"]
        reason = item.get("reason")
        if reason == "unreadable":
            lines.append(f"  {fname} — 文件无法读取")
        else:
            missing = item.get("missing", [])
            miss_str = "、".join(missing) if missing else "必要字段"
            lines.append(f"  {fname} — 缺少：{miss_str}")
    lines.append("")
    lines.append("请返回字段设置，为这些文件指定正确的列含义后重新汇总。")
    return "\n".join(lines)


# ============================================================
#  界面样式（v1.5.4 第一轮视觉改版）
#  Apple HIG 启发的浅色配色与组件尺寸；纯界面样式，不含业务逻辑。
#  仅用于欢迎页与汇总任务页，不改动其他对话框的既有外观。
# ============================================================

# 设计令牌：浅色底 / 白色卡片 / 单一克制的强调色
# Re-export the existing GUI helpers for compatible imports.
from . import gui_style as _gui_style
from .gui_style import (
    UI_BG,
    UI_SURFACE,
    UI_BORDER,
    UI_TEXT,
    UI_TEXT_SECONDARY,
    UI_TEXT_TERTIARY,
    UI_ACCENT,
    UI_ACCENT_ACTIVE,
    UI_ACCENT_PRESSED,
    UI_ACCENT_DISABLED,
    UI_BUTTON_SECONDARY,
    UI_BUTTON_SECONDARY_ACTIVE,
    UI_BUTTON_SECONDARY_PRESSED,
    UI_DISABLED_TEXT,
    UI_FIELD_BG,
    UI_OK,
    UI_WARN,
    UI_ERROR,
    UI_FONT_FAMILY,
    UI_FONT_CODE,
    MIN_WINDOW_WIDTH,
    MIN_WINDOW_HEIGHT,
    DEFAULT_WINDOW_WIDTH,
    DEFAULT_WINDOW_HEIGHT,
    DEFAULT_WINDOW_SIZE,
    COMPACT_LAYOUT_WIDTH,
    _UI_FONTS,
    _UI_FONT_SPEC,
    _UI_FONT_ROLE,
    _TkPixel,
    _fw,
    _ui_font_family,
    _ui_font,
    _ui_fonts,
    _init_ui_style,
    _ui_label,
    _ui_card,
    _ui_rule,
    _ui_badge,
    _ui_card_header,
    _ui_card_body,
    _AccentButton,
    _ui_primary_button,
    _ui_secondary_button,
    _ui_quiet_button,
    _ui_radio,
    _ui_text_field,
)


def __getattr__(name):
    # Interpreter-specific caches must be read from their single owning module.
    if name in ("_UI_STYLE_ROOT", "_UI_FONTS_ROOT"):
        return getattr(_gui_style, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


# ============================================================
#  v1.3 新增：字段选择窗口
# ============================================================

# 下拉框选项（用户可见名称 → 内部标识）
def _get_field_type_options():
    """动态生成字段类型选项列表（v1.5.0）。

    从当前字段配置读取，追加「忽略此列」选项。

    Returns:
        list[tuple[str, str]]: [(显示名, 内部名), ...]
    """
    config = get_field_config()
    options = [(f["name"], f["name"]) for f in config]
    options.append(("忽略此列", "ignore"))
    return options


def _get_field_examples():
    """动态生成字段示例说明（v1.5.0）。

    从当前字段配置读取关键词生成。
    """
    config = get_field_config()
    examples = {}
    for f in config:
        kws = f.get("keywords", [])
        examples[f["name"]] = "例如：" + "、".join(kws[:4])
    examples["ignore"] = "不参与汇总"
    return examples


class FieldMappingDialog(tk.Toplevel):
    """字段选择对话框 — 用户手动指定每列的含义。

    使用方式（外部调用者）:
        dialog = FieldMappingDialog(parent, file_path, base_dir)
        parent.wait_window(dialog)
        if dialog.result:
            # 用户点击了"确定"，映射已保存（如勾选了"记住"）
            ...

    本类不包含识别逻辑，所有数据获取和保存通过 field_mapper 接口。
    """

    def __init__(self, parent, file_path, base_dir):
        super().__init__(parent)
        self.file_path = file_path
        self.base_dir = base_dir
        self.result = False                     # True=确定, False=取消/关闭
        # v1.5.4: 本次对话框确认的映射 {表头: 字段名}。无论是否勾选「记住」，
        # 都会被外部读取并应用到当前任务（临时生效）。取消时保持为空。
        self.mappings = {}
        self._header_rows = []                  # [(col_idx, header_text), ...]
        self._combobox_rows = []                # [(col_idx, header_text, var, combo), ...]

        # 动态字段选项（v1.5.0：从配置读取）
        self._field_options = _get_field_type_options()
        self._field_examples = _get_field_examples()

        # ---------- 窗口基本设置 ----------
        self.title("请选择每一列的含义")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()

        # ---------- 读取表头 ----------
        ok, data = get_file_headers(file_path)
        if not ok:
            messagebox.showerror("无法读取文件", data, parent=self)
            self.destroy()
            return
        self._header_rows = data  # [(col_idx, header_text), ...]

        # ---------- 预选：用 detect_columns 获取部分识别结果 ----------
        preselected = self._preselect()

        # ---------- 构建界面 ----------
        self._build_ui(preselected)

        # ---------- 居中 ----------
        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w, h = self.winfo_width(), self.winfo_height()
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        self.geometry(f"+{x}+{y}")

        # 强制最小宽度（确保例文可见）
        self.minsize(540, 200)

    # ============================================================
    #  预选逻辑（v1.5.0：使用 detect_columns 返回 FieldDetection 列表）
    # ============================================================

    def _preselect(self):
        """用 _match_all_fields 获取部分识别结果，预填下拉框。

        支持 .xlsx / .xlsm（openpyxl）和 .xls（xlrd）。

        v1.5.0: 使用 _match_all_fields 直接获取部分匹配，不要求全部字段识别。

        Returns:
            dict: {col_idx: field_name, ...}
                  只包含已识别到的部分（可能什么都没有）
        """
        ext = os.path.splitext(self.file_path)[1].lower()

        # .xls — 使用 xlrd + _detect_columns_xls
        if ext == ".xls":
            try:
                import xlrd
                from .excel_reader import _detect_columns_xls
                wb = xlrd.open_workbook(self.file_path)
                try:
                    sheet = wb.sheet_by_index(0)
                    detections = _detect_columns_xls(sheet)
                    if detections is None:
                        return {}
                    return {d.col_index: d.field_name for d in detections}
                finally:
                    wb.release_resources()
            except Exception:
                return {}

        # .xlsx / .xlsm — 使用 _match_all_fields 获取部分结果
        try:
            from openpyxl import load_workbook
            from .field_mapper import _match_all_fields, _get_effective_config
            wb = load_workbook(self.file_path, data_only=True)
            try:
                ws = wb.active
                headers = [str(cell.value).strip() if cell.value is not None else ""
                           for cell in ws[1]]
            finally:
                wb.close()
            if all(h == "" for h in headers):
                return {}
            config = _get_effective_config()
            detections = _match_all_fields(headers, config)
            return {d.col_index: d.field_name for d in detections}
        except Exception:
            return {}

    # ============================================================
    #  界面构建
    # ============================================================

    def _build_ui(self, preselected):
        """构建窗口全部组件（v1.5.0 动态字段版）。"""
        # ---- 顶部说明 ----
        intro = ttk.Label(
            self,
            text=(
                "我发现您的 Excel 列名比较特殊。\n"
                "请为每一列选择正确的含义。\n"
                "这一步只需设置一次，下次会自动记住。"
            ),
            justify="left",
            padding=(16, 12),
        )
        intro.pack(fill="x")

        # ---- 分隔线 ----
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12)

        # ---- 主区域：表头 → 含义映射 ----
        # 外框 + Canvas 滚动（列多时可用）
        outer = ttk.Frame(self, padding=(12, 8))
        outer.pack(fill="both", expand=True)

        # 表头行：列名 | 含义 | 说明
        hdr_frame = ttk.Frame(outer)
        hdr_frame.pack(fill="x", pady=(0, 4))
        ttk.Label(hdr_frame, text="列名", width=16, anchor="w",
                  font=("Microsoft YaHei", 9, "bold")).pack(side="left", padx=(0, 8))
        ttk.Label(hdr_frame, text="含义", width=14, anchor="w",
                  font=("Microsoft YaHei", 9, "bold")).pack(side="left", padx=(0, 8))
        ttk.Label(hdr_frame, text="说明", anchor="w",
                  font=("Microsoft YaHei", 9, "bold")).pack(side="left")

        # 数据行
        rows_frame = ttk.Frame(outer)
        rows_frame.pack(fill="both", expand=True)

        option_display_names = [opt[0] for opt in self._field_options]

        for col_idx, header_text in self._header_rows:
            row_frame = ttk.Frame(rows_frame)
            row_frame.pack(fill="x", pady=1)

            # 列名（只读标签）
            name_label = ttk.Label(row_frame, text=header_text, width=16, anchor="w")
            name_label.pack(side="left", padx=(0, 8))

            # 下拉框（v1.5.0 动态选项）
            var = tk.StringVar()
            combo = ttk.Combobox(
                row_frame, textvariable=var,
                values=option_display_names,
                state="readonly", width=12,
            )
            # 预选
            pre_type = preselected.get(col_idx)
            if pre_type:
                for display_name, internal_name in self._field_options:
                    if internal_name == pre_type:
                        var.set(display_name)
                        break
            if not var.get():
                var.set("忽略此列")

            combo.pack(side="left", padx=(0, 8))

            # 说明文本（跟随下拉框变化）
            example_var = tk.StringVar()
            self._update_example(var.get(), example_var)
            example_label = ttk.Label(
                row_frame, textvariable=example_var,
                foreground="#888", anchor="w",
            )
            example_label.pack(side="left")

            # 绑定变化事件：更新说明文本
            var.trace_add("write",
                          lambda *_a, v=var, ev=example_var: self._update_example(v.get(), ev))

            self._combobox_rows.append((col_idx, header_text, var, combo))

        # ---- 底部：记住选择 ----
        bottom = ttk.Frame(self, padding=(12, 8))
        bottom.pack(fill="x")

        self.remember_var = tk.BooleanVar(value=True)
        cb = ttk.Checkbutton(
            bottom, text="记住我的选择（推荐）",
            variable=self.remember_var,
        )
        cb.pack(anchor="w")
        ttk.Label(
            bottom,
            text="下次遇到类似文件会自动识别。",
            foreground="#888",
            padding=(24, 0),
        ).pack(anchor="w")

        # ---- 按钮 ----
        btn_frame = ttk.Frame(self, padding=(12, 8))
        btn_frame.pack(fill="x")

        ttk.Button(btn_frame, text="确定", command=self._on_confirm).pack(
            side="right", ipadx=16, padx=(8, 0))
        ttk.Button(btn_frame, text="取消", command=self._on_cancel).pack(
            side="right", ipadx=16)

    # ============================================================
    #  事件处理（v1.5.0 动态字段版）
    # ============================================================

    def _update_example(self, display_name, example_var):
        """根据下拉框选中项刷新说明文本。"""
        for dname, iname in self._field_options:
            if dname == display_name:
                example_var.set(self._field_examples.get(iname, ""))
                return
        example_var.set("")

    def _on_confirm(self):
        """确定按钮：验证 → 保存 → 关闭（v1.5.0 动态字段版）。

        要求至少选择前两个字段（分组键 + 求和字段）。
        """
        # 收集选中的字段映射
        selected = {}  # field_name -> (col_idx, header_text)
        for col_idx, header_text, var, _combo in self._combobox_rows:
            display_name = var.get()
            for dn, inn in self._field_options:
                if dn == display_name:
                    if inn != "ignore":
                        selected[inn] = (col_idx, header_text)
                    break

        # ---- 验证：至少需要前两个字段 ----
        config = get_field_config()
        required_names = [f["name"] for f in config[:2]]

        missing = [n for n in required_names if n not in selected]
        if len(missing) == len(required_names):
            messagebox.showwarning(
                "提示",
                "您还有必要的信息没有设置，请完成后继续。",
                parent=self,
            )
            return

        for name in missing:
            messagebox.showwarning(
                "提示",
                f"未找到「{name}」，请选择正确的列。",
                parent=self,
            )
            return

        # ---- 当前任务映射（v1.5.4）----
        # 无论是否勾选「记住」，都把本次选择转换为 {表头: 字段名}，
        # 供外部应用到当前任务。勾选「记住」时才额外持久化到 settings.json。
        self.mappings = {
            header_text: field_name
            for field_name, (col_idx, header_text) in selected.items()
        }

        # ---- 持久化（仅勾选「记住」时）----
        if self.remember_var.get():
            for col_idx, header_text, var, _combo in self._combobox_rows:
                display_name = var.get()
                for dn, inn in self._field_options:
                    if dn == display_name and inn != "ignore":
                        save_field_mapping(self.base_dir, header_text, inn)
                        break

        self.result = True
        self.destroy()

    def _on_cancel(self):
        """取消按钮。"""
        self.result = False
        self.destroy()


# ============================================================
#  v1.5.2 Hotfix3 新增：Excel 文件选择窗口
# ============================================================

class ExcelFileSelectDialog(tk.Toplevel):
    """文件选择对话框 — 列出文件夹中的 Excel 文件供用户勾选。

    使用方式（外部调用者）:
        dialog = ExcelFileSelectDialog(parent, folder, excel_files)
        parent.wait_window(dialog)
        if dialog.result is not None:
            selected_names = dialog.result   # list[str]，用户勾选的文件名
        else:
            # 用户点击了取消
            ...

    默认全部勾选（保留旧「汇总整个文件夹」行为）。
    """

    def __init__(self, parent, folder, excel_files):
        super().__init__(parent)
        self.folder = folder
        self.excel_files = list(excel_files)   # 文件名列表（仅名称）
        self.result = None                     # 确定=选中名称列表, 取消/关闭=None

        # 每个文件一个勾选变量，默认全部勾选
        self._vars = {fname: tk.BooleanVar(value=True) for fname in self.excel_files}

        self.title("选择要汇总的 Excel 文件")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.minsize(420, 320)

        self._build_ui()
        self._refresh_counter()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

    def _build_ui(self):
        """构建窗口全部组件。"""
        ttk.Label(
            self, text=f"文件夹：{self.folder}",
            foreground="#888", wraplength=560, justify="left",
            padding=(12, 8),
        ).pack(fill="x")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12)

        # ---- 文件勾选列表（Canvas + 滚动） ----
        list_outer = ttk.Frame(self, padding=(12, 8))
        list_outer.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(list_outer, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(list_outer, orient="vertical", command=self.canvas.yview)
        self.list_frame = ttk.Frame(self.canvas)
        self.list_frame.bind(
            "<Configure>",
            lambda _e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.create_window((0, 0), window=self.list_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for fname in self.excel_files:
            ttk.Checkbutton(
                self.list_frame, text=fname, variable=self._vars[fname],
                command=self._refresh_counter,
            ).pack(anchor="w", pady=1)

        # ---- 批量操作：全选 / 全不选 / 反选 ----
        op_frame = ttk.Frame(self, padding=(12, 4))
        op_frame.pack(fill="x")
        ttk.Button(op_frame, text="全选", command=self._select_all).pack(side="left", padx=(0, 6))
        ttk.Button(op_frame, text="全不选", command=self._select_none).pack(side="left", padx=6)
        ttk.Button(op_frame, text="反选", command=self._invert).pack(side="left", padx=6)

        # ---- 计数器 + 确定/取消 ----
        bottom = ttk.Frame(self, padding=(12, 8))
        bottom.pack(fill="x")

        self.counter_label = ttk.Label(bottom, text="", foreground="#666")
        self.counter_label.pack(side="left")

        ttk.Button(bottom, text="确定", command=self._on_confirm).pack(
            side="right", ipadx=16, padx=(8, 0))
        ttk.Button(bottom, text="取消", command=self._on_cancel).pack(
            side="right", ipadx=16)

    # ============================================================
    #  事件处理
    # ============================================================

    def _selected_names(self):
        """返回当前勾选的文件名列表。"""
        return [fname for fname in self.excel_files if self._vars[fname].get()]

    def _refresh_counter(self):
        """刷新「已选择：N / M」计数器。"""
        n = len(self._selected_names())
        m = len(self.excel_files)
        self.counter_label.config(text=f"已选择：{n} / {m}")

    def _select_all(self):
        for var in self._vars.values():
            var.set(True)
        self._refresh_counter()

    def _select_none(self):
        for var in self._vars.values():
            var.set(False)
        self._refresh_counter()

    def _invert(self):
        for var in self._vars.values():
            var.set(not var.get())
        self._refresh_counter()

    def _on_confirm(self):
        """确定按钮：至少选择一个文件，否则警告。"""
        selected = self._selected_names()
        if not selected:
            messagebox.showwarning("提示", "请至少选择一个 Excel 文件。", parent=self)
            return
        self.result = selected
        self.destroy()

    def _on_cancel(self):
        """取消按钮。"""
        self.result = None
        self.destroy()


# ============================================================
#  v1.4.1 新增：字段规则管理窗口
# ============================================================

class FieldConfigDialog(tk.Toplevel):
    """字段规则管理对话框（v1.5.0 列表格式版）。

    使用列表 [{"name":..., "keywords":[...]}, ...] 管理字段配置。
    支持编辑字段名称、关键词，新增/删除字段，前两个字段受保护不可删除。

    功能：
        - 查看所有已配置的字段及其关键词（字段名可编辑）
        - 新增字段（输入字段名 + 逗号分隔的关键词）
        - 编辑现有字段的名称和关键词
        - 删除字段（前两个字段不允许删除）
        - 保存到 field_config.json
    """

    def __init__(self, parent, base_dir):
        super().__init__(parent)
        self.base_dir = base_dir
        self.result = False

        # 加载当前配置（列表格式：[{name, keywords}, ...]）
        self._config = get_field_config()

        # ---- 窗口设置 ----
        self.title("字段规则管理")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.minsize(560, 380)

        # ---- 构建界面 ----
        self._build_ui()

        # ---- 居中 ----
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w = self.winfo_width()
        h = self.winfo_height()
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        self.geometry(f"+{x}+{y}")

    # ============================================================
    #  界面构建
    # ============================================================

    def _build_ui(self):
        """构建窗口全部组件。"""
        # ---- 顶部说明 ----
        intro = ttk.Label(
            self,
            text=(
                "在这里可以自定义字段匹配规则。\n"
                "您可以为「客户」和「金额」添加更多关键词，\n"
                "也可以定义其他目标字段。修改后立即生效。"
            ),
            justify="left",
            padding=(16, 12),
        )
        intro.pack(fill="x")

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12)

        # ---- 字段列表区域 ----
        list_frame = ttk.Frame(self, padding=(12, 8))
        list_frame.pack(fill="both", expand=True)

        # 表头
        hdr = ttk.Frame(list_frame)
        hdr.pack(fill="x", pady=(0, 4))
        ttk.Label(hdr, text="字段名称", width=14, anchor="w",
                  font=("Microsoft YaHei", 9, "bold")).pack(side="left", padx=(0, 8))
        ttk.Label(hdr, text="匹配关键词（逗号分隔）", anchor="w",
                  font=("Microsoft YaHei", 9, "bold")).pack(side="left", padx=(0, 8))
        ttk.Label(hdr, text="操作", width=10, anchor="w",
                  font=("Microsoft YaHei", 9, "bold")).pack(side="left")

        # 可滚动字段列表
        canvas_frame = ttk.Frame(list_frame)
        canvas_frame.pack(fill="both", expand=True)

        self._rows_frame = ttk.Frame(canvas_frame)
        self._rows_frame.pack(fill="both", expand=True)

        self._refresh_field_rows()

        # ---- 新增字段区域 ----
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12)

        add_frame = ttk.Frame(self, padding=(12, 8))
        add_frame.pack(fill="x")

        ttk.Label(add_frame, text="新增字段：", font=("Microsoft YaHei", 9, "bold")).pack(
            side="left", padx=(0, 6))

        self.new_field_name = tk.StringVar()
        name_entry = ttk.Entry(add_frame, textvariable=self.new_field_name, width=12)
        name_entry.pack(side="left", padx=(0, 6))

        ttk.Label(add_frame, text="关键词：").pack(side="left", padx=(0, 4))

        self.new_field_kw = tk.StringVar()
        kw_entry = ttk.Entry(add_frame, textvariable=self.new_field_kw, width=28)
        kw_entry.pack(side="left", padx=(0, 6))

        ttk.Label(add_frame, text="（逗号分隔）", foreground="#888").pack(
            side="left", padx=(0, 8))

        ttk.Button(add_frame, text="添加", command=self._on_add_field).pack(side="left")

        # ---- 底部按钮 ----
        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12)

        bottom = ttk.Frame(self, padding=(12, 8))
        bottom.pack(fill="x")

        ttk.Button(bottom, text="恢复默认", command=self._on_reset).pack(side="left")
        ttk.Button(bottom, text="保存", command=self._on_save).pack(
            side="right", ipadx=16, padx=(8, 0))
        ttk.Button(bottom, text="取消", command=self._on_cancel).pack(
            side="right", ipadx=16)

    # ============================================================
    #  字段行管理（v1.5.0 列表格式）
    # ============================================================

    def _refresh_field_rows(self):
        """重新绘制字段行列表（基于列表格式 [{name, keywords}, ...]）。"""
        for w in self._rows_frame.winfo_children():
            w.destroy()

        # _row_vars: index → (name_var, kw_var)
        self._row_vars = {}

        for idx, field_def in enumerate(self._config):
            row = ttk.Frame(self._rows_frame)
            row.pack(fill="x", pady=1)

            # 字段名编辑框
            name_var = tk.StringVar(value=field_def["name"])
            name_entry = ttk.Entry(row, textvariable=name_var, width=14)
            name_entry.pack(side="left", padx=(0, 8))

            # 关键词编辑框
            kw_text = ", ".join(field_def.get("keywords", []))
            kw_var = tk.StringVar(value=kw_text)
            kw_entry = ttk.Entry(row, textvariable=kw_var, width=40)
            kw_entry.pack(side="left", padx=(0, 8))

            self._row_vars[idx] = (name_var, kw_var)

            # 操作按钮
            btn_frame = ttk.Frame(row)
            btn_frame.pack(side="left")

            # 前两个字段（客户、金额）不允许删除
            if idx < 2:
                ttk.Label(btn_frame, text="（必需）", foreground="#888").pack(side="left")
            else:
                ttk.Button(
                    btn_frame, text="删除", width=5,
                    command=lambda i=idx: self._on_delete_field(i),
                ).pack(side="left")

    # ============================================================
    #  事件处理（v1.5.0 列表格式）
    # ============================================================

    def _on_add_field(self):
        """添加新字段（列表格式）。"""
        name = self.new_field_name.get().strip()
        kw_str = self.new_field_kw.get().strip()

        if not name:
            messagebox.showwarning("提示", "请输入字段名称", parent=self)
            return
        if not kw_str:
            messagebox.showwarning("提示", "请输入至少一个关键词", parent=self)
            return

        # 检查字段名是否重复
        existing_names = [f["name"] for f in self._config]
        if name in existing_names:
            messagebox.showwarning("提示", f"字段「{name}」已存在，请直接编辑", parent=self)
            return

        # 解析关键词
        keywords = [k.strip() for k in kw_str.split(",") if k.strip()]
        if not keywords:
            messagebox.showwarning("提示", "请输入有效的关键词", parent=self)
            return

        self._config.append({"name": name, "keywords": keywords})
        self.new_field_name.set("")
        self.new_field_kw.set("")
        self._refresh_field_rows()

    def _on_delete_field(self, idx):
        """删除字段（保护前两个字段）。"""
        if idx < 2:
            return
        field_name = self._config[idx]["name"]
        if not messagebox.askyesno(
            "确认删除", f"确定删除字段「{field_name}」吗？", parent=self,
        ):
            return
        del self._config[idx]
        self._refresh_field_rows()

    def _on_reset(self):
        """恢复默认配置。"""
        if not messagebox.askyesno(
            "确认恢复",
            "将删除所有自定义字段规则，恢复为默认配置。\n确定继续吗？",
            parent=self,
        ):
            return
        reset_field_config(self.base_dir)
        self._config = get_field_config()
        self._refresh_field_rows()

    def _on_save(self):
        """保存配置（从编辑框读取最新数据，输出列表格式）。"""
        for idx, (name_var, kw_var) in self._row_vars.items():
            name = name_var.get().strip()
            kw_text = kw_var.get().strip()
            keywords = [k.strip() for k in kw_text.split(",") if k.strip()]

            if not name:
                messagebox.showwarning(
                    "提示",
                    f"第 {idx + 1} 个字段的名称不能为空",
                    parent=self,
                )
                return
            if not keywords:
                messagebox.showwarning(
                    "提示",
                    f"字段「{name}」的关键词不能为空",
                    parent=self,
                )
                return

            self._config[idx] = {"name": name, "keywords": keywords}

        # 保存为列表格式
        save_field_config(self.base_dir, self._config)
        self.result = True
        self.destroy()

    def _on_cancel(self):
        """取消。"""
        self.result = False
        self.destroy()


# ============================================================
#  GUI 应用
# ============================================================


class SummaryApp:
    """Excel智能汇总助手 — GUI 主程序"""

    def __init__(self, root, *, host=None):
        self.root = root
        self._layout_parent = host if host is not None else root
        if host is None:
            self.root.title(APP_FULL_NAME)
        config_dir = get_config_dir()

        # 状态变量 — 必须先于任何使用 self.<var> 的方法调用初始化
        # （v1.5.2 Hotfix1 修复：_load_output_setting 依赖 self.custom_output_dir）
        self.selected_folder = tk.StringVar()
        self.output_mode = tk.StringVar(value="default")  # "default" / "input" / "custom"
        self.custom_output_dir = tk.StringVar()
        self.task_running = False
        self._task_history_written = False
        self.worker = None
        self.output_path = ""

        # v1.3: 启动时加载用户保存的字段映射
        load_settings(config_dir)
        # v1.4.1: 加载用户自定义字段规则
        load_field_config(config_dir)
        # v1.5.1: 加载自定义输出目录；v1.5.3 起同时加载输出模式
        self._load_output_setting(config_dir)
        # v1.5.4 视觉改版：浅色底色 + 本页 ttk 样式（不改其他对话框外观）
        _init_ui_style(self.root)
        try:
            self.root.configure(background=UI_BG)
        except tk.TclError:
            pass
        # 保持原默认尺寸与原最小尺寸；内容更高的页面由内部滚动区域保证可达，
        # 仅在屏幕逻辑可用区域确实装不下时才按可用区域收敛。
        if host is None:
            self.root.geometry(self._fit_initial_geometry())
            self.root.minsize(MIN_WINDOW_WIDTH, MIN_WINDOW_HEIGHT)
            self.root.resizable(True, True)

        # 文件预览缓存
        self._preview_ok = []    # [(filename,), ...]
        self._preview_warn = []  # [(filename, headers), ...]

        # v1.5.2 Hotfix3: 用户勾选、参与汇总的文件名列表（仅名称）
        self.selected_excel_files = []

        # v1.5.2: 当前任务的记录上下文
        self._task_input_folder = ""
        self._task_input_files = []
        self._task_field_names = []

        # v1.5.4: 当前任务的人工字段映射 {文件名: {表头: 字段名}}。
        # 仅本次任务生效，优先级高于持久化映射；任务结束/新建任务时清空。
        self.task_field_mappings = {}

        self._build_welcome()
        self._build_task_frame()

        # 滚动区域占据根窗口，两个页面在其中切换
        self._scroll_host.pack(fill="both", expand=True)

        # 默认显示欢迎页
        self._show_welcome()

    # ============================================================
    #  欢迎页
    # ============================================================

    def _fit_initial_geometry(self):
        """计算初始窗口几何：默认 680×580，仅在屏幕装不下时按可用区域收敛。

        目标场景是常见的 1366×768 @125%：该屏幕的逻辑可用区域约为
        1093×614 逻辑像素（已扣除标题栏与任务栏），680×580 客户区加上窗口
        边框（实测 18×47 逻辑像素）后为 687×608，仍可完整显示。
        若所在屏幕更小（或缩放更高），这里把客户区收敛到可用区域内，
        避免默认窗口本身超出屏幕；收敛后任务页由内部滚动区域保证可达。
        """
        try:
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()
            self.root.update_idletasks()
            # 装饰尺寸 = 当前外框与客户区的差值；未映射时用系统度量回退
            frame_x = self.root.winfo_rootx() - self.root.winfo_x()
            frame_y = self.root.winfo_rooty() - self.root.winfo_y()
            border = max(0, int(frame_x)) * 2
            titlebar = max(0, int(frame_y))
            if border <= 0 or titlebar <= 0:
                import ctypes
                user32 = ctypes.windll.user32
                border = user32.GetSystemMetrics(32) * 2      # SM_CXFRAME（可缩放边框）
                border += user32.GetSystemMetrics(92) * 2     # SM_CXPADDEDBORDER
                titlebar = user32.GetSystemMetrics(4)         # SM_CYCAPTION
                titlebar += user32.GetSystemMetrics(92)
            if titlebar <= 0:
                titlebar = 31                                 # 常见 125% 缩放下的标题栏
            # 任务栏/其他 appbar 占用的高度
            try:
                import ctypes
                from ctypes import wintypes

                class _RECT(ctypes.Structure):
                    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                                ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

                rect = _RECT()
                ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
                work_h = rect.bottom - rect.top
                taskbar = max(0, screen_h - work_h)
            except Exception:
                taskbar = 40

            avail_w = screen_w - border
            avail_h = screen_h - titlebar - taskbar
            width = min(DEFAULT_WINDOW_WIDTH, max(MIN_WINDOW_WIDTH, avail_w))
            height = min(DEFAULT_WINDOW_HEIGHT, max(MIN_WINDOW_HEIGHT, avail_h))
            if (width, height) != (DEFAULT_WINDOW_WIDTH, DEFAULT_WINDOW_HEIGHT):
                print(f"[UI] 屏幕可用区域 {avail_w}×{avail_h} 小于默认窗口 "
                      f"{DEFAULT_WINDOW_WIDTH}×{DEFAULT_WINDOW_HEIGHT}，"
                      f"初始窗口收敛为 {width}×{height}")
            x = max(0, (screen_w - width - border) // 2)
            y = max(0, (screen_h - taskbar - height - titlebar) // 2)
            return f"{width}x{height}+{x}+{y}"
        except Exception:
            return DEFAULT_WINDOW_SIZE

    # ---- 滚动区域：内容高于客户区时才出现滚动条 ----

    def _create_scroll_host(self):
        """创建承载两页的滚动区域，返回放内容的 frame。

        实现说明（重要）：这里刻意不使用 tk.Canvas。
        实测在本机 Tk 8.6 / vista 主题下，放进 Canvas 窗口（create_window）
        的 ttk 控件会丢失自身背景——Card.TFrame 不再画白色、tk.Label 的
        background 也被忽略，于是卡片变成系统灰、文字出现异色矩形底。
        因此改用「place 实现的裁剪容器」：内容在一个普通 ttk.Frame 内正常
        绘制，只通过改变它的 y 偏移来滚动。

        内容自然高度小于客户区时不显示滚动条（视觉上与整页一致）；
        窗口缩到最小尺寸 540×480 时出现滚动条，保证所有控件仍可达。
        """
        host = ttk.Frame(self._layout_parent, style="Page.TFrame")
        self._scroll_host = host
        host.columnconfigure(0, weight=1)
        host.rowconfigure(0, weight=1)

        content = ttk.Frame(host, style="Page.TFrame")
        content.columnconfigure(0, weight=1)
        self._scroll_inner = content
        self._scroll_content = content
        self._scroll_offset = 0
        self._scroll_span = 0

        self._scrollbar = ttk.Scrollbar(host, orient="vertical",
                                        command=self._on_scrollbar_command)
        host.bind("<Configure>", self._on_scroll_viewport_configure)
        content.bind("<Configure>", self._on_scroll_content_configure)
        for widget in (host, content):
            widget.bind("<MouseWheel>", self._on_scroll_wheel)
        self._wheel_bound = set()
        self._layout_scroll_content()
        return content

    def _bind_page_wheel(self, *widgets):
        """把滚轮绑定挂到页面容器及其后代，保证鼠标停在卡片上也能滚动页面。

        只挂 <MouseWheel> 绑定、不改任何控件配置；用 add="+" 避免覆盖既有绑定，
        并用 _wheel_bound 去重。文件列表自身的滚轮由 _on_file_list_wheel 处理：
        列表能滚就滚列表，不能滚就返回 None 让事件继续冒泡到这里滚动页面。
        """
        queue = list(widgets)
        while queue:
            widget = queue.pop()
            if str(widget) not in self._wheel_bound:
                self._wheel_bound.add(str(widget))
                widget.bind("<MouseWheel>", self._on_scroll_wheel, add="+")
            queue.extend(widget.winfo_children())

    def _layout_scroll_content(self):
        """把内容贴到视口左上角（宽度跟随视口，高度用自然高度）。"""
        content = self._scroll_content
        width = max(0, self._scroll_host.winfo_width())
        height = max(1, content.winfo_reqheight())
        content.place(x=0, y=-self._scroll_offset, width=width, height=height)

    def _scroll_view_span(self):
        """返回 (content_height, viewport_height, max_offset)。"""
        content_h = max(1, self._scroll_content.winfo_reqheight())
        view_h = max(1, self._scroll_host.winfo_height())
        return content_h, view_h, max(0, content_h - view_h)

    def _on_scroll_viewport_configure(self, _event=None):
        self._layout_scroll_content()
        self._refresh_scrollbar()

    def _on_scroll_content_configure(self, _event=None):
        self._layout_scroll_content()
        self._refresh_scrollbar()

    def _refresh_scrollbar(self):
        """内容放得下时隐藏滚动条，否则按偏移比例显示。

        注意执行顺序：先把偏移复位并重新布局，再隐藏滚动条。
        早期写法是「先置 0、再判断 self._scroll_offset != 0」——判断恒为假，
        于是窗口从「需要滚动」跨到「无需滚动」时内容仍停在旧的负 y 偏移，
        滚动条却已隐藏（顶部标题被推到视口外）。
        """
        content_h, view_h, max_off = self._scroll_view_span()
        if max_off <= 0:
            previous = self._scroll_offset
            self._scroll_offset = 0
            self._scrollbar.grid_remove()
            self._scrollbar.set(0.0, 1.0)
            if previous != 0:
                self._layout_scroll_content()
            return
        self._scroll_offset = min(self._scroll_offset, max_off)
        self._scrollbar.grid(row=0, column=1, sticky="ns")
        first = self._scroll_offset / content_h
        last = min(1.0, (self._scroll_offset + view_h) / content_h)
        self._scrollbar.set(first, last)

    def _on_scrollbar_command(self, *args):
        """滚动条拖动/点按：'moveto' 或 'scroll n units/pages'。"""
        content_h, view_h, max_off = self._scroll_view_span()
        if max_off <= 0:
            return
        if args and args[0] == "moveto":
            self._scroll_offset = int(float(args[1]) * content_h)
        elif args and args[0] == "scroll":
            amount = int(args[1])
            unit = 20
            if len(args) > 2 and args[2] == "pages":
                unit = max(1, view_h - 20)
            self._scroll_offset += amount * unit
        self._scroll_offset = max(0, min(self._scroll_offset, max_off))
        self._layout_scroll_content()
        self._refresh_scrollbar()

    def _on_scroll_wheel(self, event):
        content_h, view_h, max_off = self._scroll_view_span()
        if max_off <= 0:
            return "break"
        self._scroll_offset = max(0, min(self._scroll_offset + int(-event.delta / 120) * 20,
                                         max_off))
        self._layout_scroll_content()
        self._refresh_scrollbar()
        return "break"

    def _refresh_scroll_region(self):
        """重新计算滚动范围与滚动条显隐（切换页面/内容变化后调用）。"""
        self.root.after_idle(self._on_scroll_content_configure)

    # ---- 窄窗口自适应：优先隐藏非关键说明文字 ----

    def _bind_responsive_layout(self):
        for page in (self.welcome_frame, self.task_frame):
            page.bind("<Configure>", self._on_page_configure)

    def _on_page_configure(self, _event=None):
        if getattr(self, "_responsive_pending", False):
            return
        self._responsive_pending = True
        self.root.after_idle(self._apply_responsive_layout)

    def _apply_responsive_layout(self):
        """客户区过窄时隐藏卡片头部/导航的辅助说明，避免标题被挤到换行或截断。"""
        self._responsive_pending = False
        width = self.root.winfo_width()
        compact = width < COMPACT_LAYOUT_WIDTH
        for label in self._responsive_hints:
            try:
                if compact:
                    label.grid_remove()
                else:
                    label.grid()
            except tk.TclError:
                pass
        if getattr(self, "_nav_hint", None) is not None:
            try:
                if compact:
                    self._nav_hint.grid_remove()
                else:
                    self._nav_hint.grid()
            except tk.TclError:
                pass
        if getattr(self, "_action_hint", None) is not None:
            try:
                if compact:
                    self._action_hint.pack_forget()
                elif not self._action_hint.winfo_ismapped():
                    self._action_hint.pack(side="right")
            except tk.TclError:
                pass

    def _build_welcome(self):
        """欢迎页：品牌标题 + 主操作 + 次要入口（浅色底、白色卡片）。"""
        inner = self._create_scroll_host()
        self.welcome_frame = ttk.Frame(inner, padding=16, style="Page.TFrame")
        self.welcome_frame.grid(row=0, column=0, sticky="nsew")
        # 卡片在窗口内水平/垂直居中；窗口太矮时保持顶部对齐并可滚动
        self.welcome_frame.rowconfigure(0, weight=1)
        self.welcome_frame.rowconfigure(2, weight=1)
        self.welcome_frame.columnconfigure(0, weight=1)

        card_border, card = _ui_card(self.welcome_frame, padding=(32, 28))
        card_border.grid(row=1, column=0, pady=(0, 2))

        # ---- 品牌区 ----
        _ui_label(card, APP_NAME, "title", UI_TEXT, UI_SURFACE).grid(row=0, column=0)
        _ui_label(card, APP_VERSION, "subtitle", UI_TEXT_SECONDARY, UI_SURFACE).grid(
            row=1, column=0, pady=(6, 0))
        _ui_label(
            card,
            "快速汇总多个 Excel 文件  ·  自动识别列名  ·  一键导出",
            "body", UI_TEXT_SECONDARY, UI_SURFACE,
            justify="center", wraplength=380,
        ).grid(row=2, column=0, pady=(14, 0))

        # ---- 主操作 ----
        start_btn = _ui_primary_button(card, "开始新的汇总任务", self._on_start_new_task)
        start_btn.grid(row=3, column=0, pady=(24, 0))

        # ---- 分隔线 + 次要入口（v1.5.2） ----
        _ui_rule(card).grid(row=4, column=0, sticky="ew", pady=(22, 14))

        secondary = ttk.Frame(card, style="Card.TFrame")
        secondary.grid(row=5, column=0)
        _ui_secondary_button(secondary, "查看任务记录", self._on_open_history).pack(
            side="left", padx=(0, 8))
        _ui_secondary_button(secondary, "关于软件", self._on_open_about).pack(side="left")
        _ui_secondary_button(card, "月度与对账工作台（开发预览）", self._on_open_workbench).grid(
            row=6, column=0, pady=(14, 0))

        # ---- 底部提示 ----
        _ui_label(
            self.welcome_frame,
            "支持 .xlsx 格式  ·  自动跳过临时文件  ·  无网络需求",
            "small", UI_TEXT_TERTIARY, UI_BG,
            justify="center", wraplength=420,
        ).grid(row=3, column=0, pady=(16, 0))

        # 鼠标停在卡片/文字上也要能滚动页面
        self._bind_page_wheel(self.welcome_frame, card_border)

    def _build_task_frame(self):
        """汇总任务页：顶部导航 + 选择文件夹 → 文件检查 → 开始汇总 + 运行状态。"""
        inner = self._scroll_inner  # 与欢迎页共用同一个滚动区域
        self.task_frame = ttk.Frame(inner, padding=20, style="Page.TFrame")
        self.task_frame.columnconfigure(0, weight=1)
        self.task_frame.rowconfigure(6, weight=1)   # 运行状态卡片吸收多余高度

        # 窄窗口时隐藏的非关键说明文字（标题/提示保留）
        self._responsive_hints = []
        self._nav_hint = None
        self._action_hint = None
        self._responsive_pending = False

        # ---- 顶部导航 ----
        nav = ttk.Frame(self.task_frame, style="Page.TFrame")
        nav.grid(row=0, column=0, sticky="ew")
        nav.columnconfigure(1, weight=1)

        back_btn = _ui_quiet_button(nav, "← 返回首页", self._show_welcome)
        back_btn.grid(row=0, column=0, sticky="w")
        _ui_label(nav, "新建汇总任务", "page_title", UI_TEXT, UI_BG).grid(
            row=0, column=1, sticky="w", padx=(12, 0))
        self._nav_hint = _ui_label(nav, "选择文件夹 → 检查文件 → 开始汇总", "small",
                                   UI_TEXT_TERTIARY, UI_BG)
        self._nav_hint.grid(row=0, column=2, sticky="e", padx=(10, 0))

        _ui_rule(self.task_frame, background=UI_BORDER).grid(
            row=1, column=0, sticky="ew", pady=(12, 12))

        # ---- 步骤 1：选择文件夹 ----
        step1_border, step1 = _ui_card(self.task_frame)
        step1_border.grid(row=2, column=0, sticky="ew")
        _ui_card_header(step1, 1, "选择文件夹", "存放 Excel 文件的目录")

        row1 = _ui_card_body(step1)
        row1.columnconfigure(0, weight=1)
        # 保留系统原生输入框外观（ttk.Entry 不支持 relief/borderwidth 选项）
        self.folder_entry = ttk.Entry(
            row1, textvariable=self.selected_folder, font=_ui_font("body"),
        )
        self.folder_entry.grid(row=0, column=0, sticky="ew", ipady=2, padx=(0, 10))
        _ui_secondary_button(row1, "浏览...", self._on_browse_folder).grid(
            row=0, column=1, sticky="e")

        # ---- 步骤 2：文件检查 ----
        step2_border, step2 = _ui_card(self.task_frame)
        step2_border.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        _ui_card_header(step2, 2, "文件检查", "确认列名是否可自动识别")

        body2 = _ui_card_body(step2)
        body2.rowconfigure(1, weight=1)

        self.file_status_label = _ui_label(
            body2, "请先选择文件夹", "body", UI_TEXT_TERTIARY, UI_SURFACE)
        self.file_status_label.grid(row=0, column=0, sticky="w")

        # 文件列表自身可滚动：视口 3 行，超过的文件（含警告行与「设置列」按钮）
        # 通过右侧滚动条可达；新增内容时自动滚到底部，保证第 4 行起不被裁切。
        list_row = ttk.Frame(body2, style="Card.TFrame")
        list_row.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        list_row.columnconfigure(0, weight=1)

        self.file_list_text = _ui_text_field(list_row, height=3, width=34,
                                            state="disabled")
        self.file_list_text.grid(row=0, column=0, sticky="ew", ipady=3)
        self._file_list_scrollbar = ttk.Scrollbar(
            list_row, orient="vertical", command=self.file_list_text.yview)
        self._file_list_scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_list_text.configure(yscrollcommand=self._on_file_list_scrolled)
        self._file_list_scrollbar.grid_remove()
        self.file_list_text.bind("<MouseWheel>", self._on_file_list_wheel)
        self._file_list_line_count = 0

        # v1.4.1: 字段规则管理入口
        footer2 = ttk.Frame(body2, style="Card.TFrame")
        footer2.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        _ui_quiet_button(footer2, "管理字段规则", self._on_open_field_config).pack(side="right")

        # ---- 步骤 3：开始汇总 ----
        step3_border, step3 = _ui_card(self.task_frame)
        step3_border.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        _ui_card_header(step3, 3, "开始汇总", "")

        body3 = _ui_card_body(step3)
        action_row = ttk.Frame(body3, style="Card.TFrame")
        action_row.grid(row=0, column=0, sticky="ew")

        # 主操作：强调色按钮（state / text 仍由原有回调控制）
        self.start_btn = _ui_primary_button(action_row, "开始汇总", self._on_start_summary)
        self.start_btn.configure(state="disabled")
        self.start_btn.pack(side="left")

        self._action_hint = _ui_label(action_row, "确认输出位置后开始汇总", "small",
                                      UI_TEXT_TERTIARY, UI_SURFACE)
        self._action_hint.pack(side="right")

        # v1.5.1: 输出位置选择（独立一行，窄窗口下也不截断）
        output_row = ttk.Frame(body3, style="Card.TFrame")
        output_row.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self._output_row = output_row

        _ui_label(output_row, "输出位置", "small", UI_TEXT_TERTIARY, UI_SURFACE).pack(
            side="left", padx=(0, 10))
        _ui_radio(output_row, "默认输出目录", self.output_mode, "default").pack(
            side="left", padx=(0, 4))
        _ui_radio(output_row, "输入文件夹", self.output_mode, "input").pack(
            side="left", padx=(0, 4))
        _ui_radio(output_row, "自定义目录", self.output_mode, "custom").pack(side="left")

        # 自定义目录选择按钮（初始隐藏，选中"自定义目录"时显示）
        self._custom_dir_btn = _ui_secondary_button(
            output_row, "选择目录", self._on_select_custom_dir, width=9)
        self.output_mode.trace_add("write", self._on_output_mode_changed)

        # ---- 运行状态 ----
        status_border, status_card = _ui_card(self.task_frame)
        status_border.grid(row=6, column=0, sticky="nsew", pady=(8, 0))
        self._status_border = status_border
        self._status_card = status_card
        status_card.rowconfigure(1, weight=1)
        status_card.columnconfigure(0, weight=1)
        _ui_label(status_card, "运行状态", "card_title", UI_TEXT, UI_SURFACE).grid(
            row=0, column=0, sticky="w")

        # 状态文本直接作为卡片第二行，运行时吸收多余高度
        self.status_text = _ui_text_field(status_card, height=3, width=34,
                                          state="disabled", wrap="word")
        self.status_text.grid(row=1, column=0, sticky="nsew", pady=(6, 0), ipady=3)

        # 进度条（运行中才显示，结束后隐藏）
        self.progress = ttk.Progressbar(status_card, mode="indeterminate")

        # ---- 结果操作 ----
        result_row = ttk.Frame(self.task_frame, style="Page.TFrame")
        result_row.grid(row=7, column=0, sticky="ew", pady=(8, 0))

        self.open_btn = _ui_secondary_button(
            result_row, "打开汇总文件", self._on_open_file)
        self.open_btn.configure(state="disabled")
        self.open_btn.pack(side="left")

        self.open_folder_btn = _ui_secondary_button(
            result_row, "打开文件位置", self._on_open_folder)
        self.open_folder_btn.configure(state="disabled")
        self.open_folder_btn.pack(side="left", padx=(8, 0))

        # v1.5.3: 同步「选择目录」按钮初始显隐。
        # 必须在此处调用：_load_output_setting() 在界面创建之前就把 output_mode
        # 设为 custom，之后注册的 trace 不会再触发，缺这一步会导致首次进入任务页
        # 时「选择目录」按钮缺失。
        self._sync_custom_dir_button()

        # 收集卡片头部的非关键说明，交给窄窗口自适应隐藏
        for card in (step1, step2, step3):
            self._responsive_hints.extend(getattr(card, "_ui_hints", []))
        self._bind_responsive_layout()
        # 鼠标停在卡片/列表上也要能滚动页面（列表能滚时优先滚列表）
        self._bind_page_wheel(self.task_frame, step1_border, step2_border, step3_border)

    # ---- 文件列表滚动（保证「设置列」按钮始终可达） ----

    def _on_file_list_scrolled(self, first, last):
        """内容超出一屏时显示列表滚动条。"""
        self._file_list_scrollbar.set(first, last)
        if float(first) <= 0.0 and float(last) >= 1.0:
            self._file_list_scrollbar.grid_remove()
        else:
            self._file_list_scrollbar.grid(row=0, column=1, sticky="ns")

    def _scroll_file_list_to_end(self):
        """新增内容后自动滚到底部：警告行与其中的「设置列」按钮不再落在视口外。"""
        self.file_list_text.see("end")

    def _on_file_list_wheel(self, event):
        """列表内滚轮：列表在该方向还能滚就滚列表，否则交给外层页面滚动。

        早期实现依赖 Tk 的事件冒泡（返回 None 让父级处理），但 Text 还有一条
        类绑定会滚动自身，实际结果不可靠；只判断「列表能否滚动」也不够——
        列表已到顶部/底部时滚轮会被白白吞掉，鼠标停在列表上就再也滚不动页面。
        这里按方向判断并显式委派，始终返回 "break" 阻止 Text 类绑定再动列表。
        """
        text = self.file_list_text
        first, last = text.yview()
        scrollable = float(last) - float(first) < 1.0
        going_down = int(-event.delta / 120) > 0
        can_move = scrollable and ((going_down and float(last) < 1.0)
                                   or (not going_down and float(first) > 0.0))
        if can_move:
            text.yview_scroll(int(-event.delta / 120), "units")
            return "break"
        # 列表放得下或已到边界：驱动页面滚动，行为与注释一致
        self._on_scroll_wheel(event)
        return "break"

    # ============================================================
    #  页面切换
    # ============================================================

    def _show_welcome(self):
        """切到欢迎页（两页共用同一个滚动区域，在其内部切换）。"""
        self.task_frame.grid_remove()
        self.welcome_frame.grid(row=0, column=0, sticky="nsew")
        self._scroll_offset = 0
        self._refresh_scroll_region()

    def _show_task(self):
        """切到任务页（两页共用同一个滚动区域）。"""
        self.welcome_frame.grid_remove()
        self.task_frame.grid(row=0, column=0, sticky="nsew")
        # 默认填入程序目录下的"对账单"文件夹
        default_folder = os.path.join(get_base_dir(), FOLDER_NAME)
        if os.path.isdir(default_folder) and not self.selected_folder.get():
            self.selected_folder.set(default_folder)
            self._on_folder_selected()
        self._scroll_offset = 0
        self._refresh_scroll_region()

    # ============================================================
    #  交互事件
    # ============================================================

    def _on_start_new_task(self):
        self.selected_folder.set("")
        # v1.5.3: 新建任务只重置输入（文件夹/文件列表），不再重置输出模式与自定义输出目录
        self._preview_ok.clear()
        self._preview_warn.clear()
        self.selected_excel_files = []
        # v1.5.4: 清空上一任务的临时字段映射，防止污染新任务
        self.task_field_mappings = {}
        self._clear_status()
        self._show_task()

    def _on_browse_folder(self):
        # 使用 Windows 原生文件夹选择器：Tkinter 的 askdirectory()
        # 在 Windows 上直接调用系统原生文件夹浏览界面。
        folder = filedialog.askdirectory(title="选择包含 Excel 文件的文件夹")
        if folder:
            self.selected_folder.set(folder)
            self._on_folder_selected()

    def _on_folder_selected(self):
        """选择文件夹后：扫描 Excel 文件 → 弹出文件勾选对话框 → 仅勾选文件参与汇总。

        v1.5.2 Hotfix3: 不再自动汇总整个文件夹，改为让用户勾选文件。
        """
        logs_dir = get_logs_dir()
        log_path = os.path.join(logs_dir, "startup.log")
        folder = self.selected_folder.get()
        self._clear_status()
        self._preview_ok.clear()
        self._preview_warn.clear()
        self.selected_excel_files = []
        # v1.5.4: 重新选择文件夹视为新任务上下文，清空上一任务的临时字段映射
        self.task_field_mappings = {}

        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"SCAN: folder={folder}\n")

        if not folder or not os.path.isdir(folder):
            self.file_status_label.config(text="请先选择文件夹", foreground=UI_TEXT_TERTIARY)
            self._update_file_list_display()
            self.start_btn.config(state="disabled")
            return

        # 扫描 Excel 文件（复用 excel_reader._get_excel_files，过滤临时/隐藏/非 Excel/空文件）
        excel_files = _get_excel_files(folder)

        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"SCAN: excel_files={excel_files}\n")

        if not excel_files:
            self.file_status_label.config(
                text="当前文件夹中没有可汇总的 Excel 文件。", foreground=UI_ERROR
            )
            self._update_file_list_display()
            self.start_btn.config(state="disabled")
            return

        # 弹出文件勾选对话框（默认全部勾选）
        selected = self._prompt_select_files(folder, sorted(excel_files))
        if selected is None:
            # 用户取消 → 未选择任何文件
            self.selected_excel_files = []
            self.file_status_label.config(text="已取消选择文件", foreground=UI_TEXT_TERTIARY)
            self._update_file_list_display()
            self.start_btn.config(state="disabled")
            return

        self.selected_excel_files = list(selected)
        self._rescan_selected_files()

    def _rescan_selected_files(self):
        """仅对 self.selected_excel_files 做字段识别检查，刷新文件列表与状态。"""
        logs_dir = get_logs_dir()
        log_path = os.path.join(logs_dir, "startup.log")
        folder = self.selected_folder.get()
        self._preview_ok.clear()
        self._preview_warn.clear()

        from openpyxl import load_workbook
        try:
            import xlrd
        except ImportError:
            xlrd = None

        for fname in sorted(self.selected_excel_files):
            fpath = os.path.join(folder, fname)
            ext = os.path.splitext(fname)[1].lower()
            # v1.5.4: 优先使用当前任务人工映射（临时生效）
            header_map = self.task_field_mappings.get(fname)
            try:
                if ext in (".xlsx", ".xlsm"):
                    wb = load_workbook(fpath)
                    try:
                        ws = wb.active
                        detections = detect_columns(ws, header_map)
                        headers = get_header_list(ws)
                    finally:
                        wb.close()
                else:  # .xls
                    wb_xls = xlrd.open_workbook(fpath)
                    try:
                        sheet = wb_xls.sheet_by_index(0)
                        headers = sheet.row_values(0)
                        detections = _detect_columns_xls(sheet, header_map)
                    finally:
                        wb_xls.release_resources()
                if detections is not None and len(detections) >= 2:
                    self._preview_ok.append(fname)
                else:
                    self._preview_warn.append((fname, headers))
            except Exception as e:
                with open(log_path, "a", encoding="utf-8") as f:
                    f.write(f"SCAN: error on {fname}: {e}\n")
                self._preview_warn.append((fname, ["（无法读取）"]))

        total = len(self._preview_ok) + len(self._preview_warn)
        ok_count = len(self._preview_ok)
        warn_count = len(self._preview_warn)

        if warn_count == 0:
            self.file_status_label.config(
                text=f"已选择 {total} 个 Excel 文件，全部可识别", foreground=UI_OK
            )
        else:
            self.file_status_label.config(
                text=f"已选择 {total} 个文件，{ok_count} 个可识别，{warn_count} 个需要检查",
                foreground=UI_WARN,
            )

        self._update_file_list_display()
        self.start_btn.config(state="normal" if ok_count > 0 else "disabled")

    def _prompt_select_files(self, folder, excel_files):
        """弹出文件勾选对话框，返回用户勾选的文件名列表；取消返回 None。"""
        dialog = ExcelFileSelectDialog(self.root, folder, excel_files)
        self.root.wait_window(dialog)
        return dialog.result

    def _update_file_list_display(self):
        """刷新文件列表显示。"""
        logs_dir = get_logs_dir()
        log_path = os.path.join(logs_dir, "startup.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"SELF_FILES_AFTER_SCAN: ok={self._preview_ok} warn_count={len(self._preview_warn)}\n")
        text_widget = self.file_list_text
        text_widget.config(state="normal")
        text_widget.delete("1.0", "end")

        for fname in self._preview_ok:
            text_widget.insert("end", f"    ✓  {fname}\n", "ok")
        for fname, headers in self._preview_warn:
            hdr_str = ", ".join(str(h) for h in headers if h is not None)
            text_widget.insert("end", f"    ⚠  {fname} — 表头 [{hdr_str}] 无法识别  ", "warn")
            # 嵌入 [设置列] 按钮
            btn = ttk.Button(
                text_widget, text="设置列", width=7,
                command=lambda fn=fname: self._on_open_mapping_dialog(fn),
            )
            text_widget.window_create("end", window=btn)
            text_widget.insert("end", "\n")

        if not self._preview_ok and not self._preview_warn:
            text_widget.insert("end", "    （未发现 Excel 文件）", "empty")

        text_widget.tag_config("ok", foreground=UI_OK)
        text_widget.tag_config("warn", foreground=UI_WARN)
        text_widget.tag_config("empty", foreground=UI_TEXT_TERTIARY)
        text_widget.config(state="disabled")
        # 记录实际行数（含警告行），并滚到底部让最后一行完整可见
        self._file_list_line_count = (len(self._preview_ok) + len(self._preview_warn)
                                      + (1 if not self._preview_ok and not self._preview_warn else 0))
        self.root.after_idle(self._scroll_file_list_to_end)

    def _on_open_mapping_dialog(self, filename):
        """打开字段选择对话框，用户映射后自动重新扫描。

        流程:
            1. 构建文件路径
            2. 弹出 FieldMappingDialog（模态）
            3. 用户确定且保存后 → 重新扫描 → 刷新文件列表
        """
        folder = self.selected_folder.get()
        file_path = os.path.join(folder, filename)

        logs_dir = get_logs_dir()
        log_path = os.path.join(logs_dir, "startup.log")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"MAPPING_FILES: folder={folder} filename={filename} file_path={file_path}\n")

        dialog = FieldMappingDialog(self.root, file_path, get_config_dir())
        self.root.wait_window(dialog)

        if dialog.result:
            # v1.5.4: 把本次对话框确认的映射应用到当前任务（临时生效）。
            # 无论是否勾选「记住」，当前任务都会使用该映射；
            # 勾选「记住」时 FieldMappingDialog 内部已写入 settings.json。
            if dialog.mappings:
                self.task_field_mappings[filename] = dict(dialog.mappings)

            # v1.5.2 Hotfix3: 仅对已勾选文件重新识别，不重新扫描整个文件夹
            if self.selected_excel_files:
                self._rescan_selected_files()
            else:
                self._on_folder_selected()

    def _on_open_field_config(self):
        """打开字段规则管理对话框。"""
        dialog = FieldConfigDialog(self.root, get_config_dir())
        self.root.wait_window(dialog)

        if dialog.result:
            # 配置已更新，重新识别当前已勾选文件（如果有）
            if self.selected_folder.get() and self.selected_excel_files:
                self._rescan_selected_files()

    # ============================================================
    #  v1.5.1: 自定义输出目录管理
    # ============================================================

    def _load_output_setting(self, config_dir):
        """从 settings.json 加载输出模式与自定义输出目录。

        v1.5.3: 新增 output_mode 持久化，与 output_dir 一并加载。
        无 settings.json / 键缺失 / 目录无效时静默保持默认。
        """
        import json
        settings_path = os.path.join(config_dir, "settings.json")
        if not os.path.isfile(settings_path):
            return
        try:
            with open(settings_path, "r", encoding="utf-8") as f:
                settings = json.load(f)
            # 输出模式（default / input / custom）
            saved_mode = settings.get("output_mode", "")
            if saved_mode in ("default", "input", "custom"):
                self.output_mode.set(saved_mode)
            # 自定义输出目录
            saved = settings.get("output_dir", "")
            if isinstance(saved, str):
                saved = saved.strip()
            else:
                saved = ""
            if saved and os.path.isdir(saved):
                self.custom_output_dir.set(saved)
        except (json.JSONDecodeError, OSError):
            pass

    def _save_output_setting(self):
        """保存输出模式与自定义输出目录到 settings.json（v1.5.3）。"""
        import json
        config_dir = get_config_dir()
        settings_path = os.path.join(config_dir, "settings.json")
        settings = {}
        if os.path.isfile(settings_path):
            try:
                with open(settings_path, "r", encoding="utf-8") as f:
                    settings = json.load(f)
            except (json.JSONDecodeError, OSError):
                settings = {}
        settings["output_mode"] = self.output_mode.get().strip()
        settings["output_dir"] = self.custom_output_dir.get().strip()
        with open(settings_path, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)

    def _on_select_custom_dir(self):
        """选择自定义输出目录。"""
        directory = filedialog.askdirectory(
            title="选择输出目录",
            parent=self.root,
        )
        if directory:
            self.custom_output_dir.set(directory)
            self._save_output_setting()

    def _sync_custom_dir_button(self):
        """根据当前输出模式显示/隐藏「选择目录」按钮（v1.5.3 拆分）。"""
        if self.output_mode.get() == "custom":
            self._custom_dir_btn.pack(side="left", padx=4)
            if not self.custom_output_dir.get().strip():
                # 默认填入程序目录下的 Excel汇总结果
                default = os.path.join(get_base_dir(), OUTPUT_FOLDER_NAME)
                self.custom_output_dir.set(default)
        else:
            self._custom_dir_btn.pack_forget()

    def _on_output_mode_changed(self, *_args):
        """输出模式切换时：同步按钮显隐，并持久化输出设置。"""
        self._sync_custom_dir_button()
        self._save_output_setting()

    def _on_start_summary(self):
        """点击 [开始汇总] — 启动后台线程。"""
        if self.task_running:
            return

        folder = self.selected_folder.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择有效的文件夹")
            return

        # v1.5.2 Hotfix3: 必须使用用户勾选的文件，禁止再次扫描整个文件夹
        if not self.selected_excel_files:
            messagebox.showwarning("提示", "请先选择要汇总的 Excel 文件")
            return

        # v1.5.1: 确定输出目录
        mode = self.output_mode.get()
        if mode == "input":
            output_dir = folder
        elif mode == "custom":
            custom = self.custom_output_dir.get().strip()
            if not custom or not os.path.isdir(custom):
                messagebox.showwarning("提示", "请先选择有效的自定义输出目录")
                return
            output_dir = custom
        else:
            output_dir = os.path.join(get_base_dir(), OUTPUT_FOLDER_NAME)
        # 防覆盖：自动生成唯一文件名
        self.output_path = get_unique_file_path(output_dir, OUTPUT_FILE)

        # v1.5.2 Hotfix3: 记录当前任务上下文（仅记录用户勾选的文件）
        self._task_input_folder = folder
        self._task_input_files = list(self.selected_excel_files)
        self._task_field_names = [f["name"] for f in get_field_config()]
        self._task_history_written = False

        # UI 进入"运行中"状态
        self.task_running = True
        self.start_btn.config(state="disabled", text="处理中...")
        self.open_btn.config(state="disabled")
        self.open_folder_btn.config(state="disabled")
        self.progress.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        self.progress.start(10)

        self._clear_status()
        self._append_status("开始汇总...", "title")

        # 启动后台线程（只汇总用户勾选的文件）
        self.worker = _WorkerThread(
            folder, self.output_path, get_logs_dir(),
            file_names=list(self.selected_excel_files),
            field_mappings=self.task_field_mappings,
        )
        self.worker.start()

        # 定时轮询队列
        self._poll_queue()

    def _poll_queue(self):
        """每 100ms 检查后台线程的消息队列。"""
        if self.worker is None:
            return

        try:
            while True:
                msg = self.worker.queue.get_nowait()
                self._handle_worker_msg(msg)
        except queue.Empty:
            pass

        if self.worker.is_alive():
            self.root.after(100, self._poll_queue)
        else:
            # 线程已结束但可能还有残留消息
            self._drain_queue()
            self.task_running = False
            self.start_btn.config(state="normal", text="重新汇总")
            self.progress.stop()
            self.progress.grid_remove()

    def _drain_queue(self):
        """取空队列中剩余消息。"""
        if self.worker is None:
            return
        try:
            while True:
                msg = self.worker.queue.get_nowait()
                self._handle_worker_msg(msg)
        except queue.Empty:
            pass

    def _handle_worker_msg(self, msg):
        """处理后台线程发来的单条消息。"""
        msg_type = msg[0]

        if msg_type == "status":
            self._append_status(msg[1], "info")

        elif msg_type == "log":
            # 原 core 模块 print 的输出
            self._append_status(f"  {msg[1]}", "log")

        elif msg_type == "complete":
            info = msg[1]
            self._append_status("", "br")
            self._append_status("汇总完成！", "title")
            self._append_status(f"  读取文件：{info['file_count']} 个", "result")
            self._append_status(f"  客户数量：{info['customer_count']} 个", "result")
            self._append_status(f"  汇总金额：{info['sum_money']:,.2f} 元", "result")
            self._append_status(f"  生成文件：{info['output_path']}", "result")
            self._append_status("", "br")

            # v1.5.2: 仅在输出文件真正生成后记录成功任务
            self._record_task(
                status="success",
                file_count=info.get("file_count", 0),
                field_names=info.get("field_names", []),
                output_path=info.get("output_path", ""),
                error="",
                result_count=info.get("customer_count", 0),
            )

            # 启用结果操作按钮
            self.open_btn.config(state="normal")
            self.open_folder_btn.config(state="normal")

        elif msg_type == "file_locked":
            self._append_status("错误：汇总文件正在打开，请关闭后重新生成。", "error")
            self.task_running = False
            self.start_btn.config(state="normal", text="开始汇总")
            self.progress.stop()
            self.progress.grid_remove()
            # v1.5.2: 记录失败任务
            self._record_task(
                status="failed",
                file_count=0,
                field_names=self._task_field_names,
                output_path="",
                error="汇总文件被占用，无法写入",
            )
            messagebox.showerror(
                "文件被占用",
                f"汇总文件正在打开，请关闭后重新生成。\n\n{msg[1]}",
            )

        elif msg_type == "error":
            self._append_status(f"错误：{msg[1]}", "error")
            self.task_running = False
            self.start_btn.config(state="normal", text="开始汇总")
            self.progress.stop()
            self.progress.grid_remove()
            # v1.5.2: 记录失败任务
            self._record_task(
                status="failed",
                file_count=0,
                field_names=self._task_field_names,
                output_path="",
                error=str(msg[1]),
            )

    # ============================================================
    #  v1.5.2: 任务记录
    # ============================================================

    def _record_task(self, status, file_count, field_names, output_path, error, *, result_count=0):
        """One terminal metadata record; legacy args kept, inputs/errors never persisted."""
        if self._task_history_written:
            return
        self._task_history_written = True
        count = file_count if status == "success" else len(self._task_input_files)
        if not task_history.save_task("regular", status, phase="summary", file_count=count,
                                      output_path=output_path, result_count=result_count):
            self._append_status("本次记录未保存，原记录文件保留；汇总结果不受影响。", "info")

    def _on_open_history(self):
        """打开任务记录窗口。"""
        HistoryDialog(self.root)

    def _on_open_workbench(self):
        """Open one modal workbench; the existing summary pipeline remains intact."""
        if self.task_running:
            messagebox.showwarning("请先完成当前任务", "原汇总任务仍在运行，请完成后再打开工作台。", parent=self.root)
            return
        window = getattr(self, "_workbench_window", None)
        if window is not None and window.winfo_exists():
            window.lift()
            return
        from .workbench_gui import WorkbenchWindow
        self._workbench_window = WorkbenchWindow(
            self.root, os.path.join(get_config_dir(), "monthly_templates.json"),
            history_writer=task_history.save_task)
        self._workbench_window.grab_set()

    def _on_open_about(self):
        """打开「关于软件」对话框。"""
        AboutDialog(self.root)

    def _on_open_file(self):
        """用系统默认程序打开汇总文件。"""
        if self.output_path and os.path.isfile(self.output_path):
            os.startfile(self.output_path)

    def _on_open_folder(self):
        """打开汇总文件所在文件夹。"""
        if self.output_path:
            folder = os.path.dirname(self.output_path)
            if os.path.isdir(folder):
                os.startfile(folder)

    # ============================================================
    #  工具方法
    # ============================================================

    def _clear_status(self):
        text = self.status_text
        text.config(state="normal")
        text.delete("1.0", "end")
        text.config(state="disabled")

    def _append_status(self, message, tag=None):
        text = self.status_text
        text.config(state="normal")
        text.insert("end", message + "\n", tag or "")
        if tag == "title":
            text.tag_config("title", font=_ui_font("card_title"), foreground=UI_TEXT)
        elif tag == "result":
            text.tag_config("result", foreground=UI_OK, font=_ui_font("body"))
        elif tag == "error":
            text.tag_config("error", foreground=UI_ERROR)
        elif tag == "log":
            text.tag_config("log", foreground=UI_TEXT_TERTIARY)
        text.see("end")
        text.config(state="disabled")


# ============================================================
#  v1.5.2: 任务记录窗口
# ============================================================

class HistoryDialog(tk.Toplevel):
    """任务记录窗口 — 展示历史汇总任务列表与详情。

    提供：查看详情（选中即显示）、打开输出文件、打开所在文件夹、
    删除单条记录、清空全部记录。所有操作只影响 history.json，
    绝不删除任何实际 Excel 文件。
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("任务记录")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.minsize(720, 500)

        self._records = load_history()

        self._build_ui()
        self._refresh_list()

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

    def _build_ui(self):
        intro = ttk.Label(
            self,
            text="本地任务记录（最多保留 500 条，最新任务在上方；预览不等于已导出）",
            padding=(12, 8),
        )
        intro.pack(fill="x")
        self._history_intro = intro

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12)

        # ---- 列表 ----
        columns = ("time", "task_type", "status", "file_count", "output_file")
        self.tree = ttk.Treeview(
            self, columns=columns, show="headings", selectmode="browse",
        )
        self.tree.heading("time", text="时间")
        self.tree.heading("task_type", text="任务类型")
        self.tree.column("task_type", width=115, anchor="w")
        self.tree.heading("status", text="状态")
        self.tree.heading("file_count", text="文件数")
        self.tree.heading("output_file", text="输出文件")
        self.tree.column("time", width=150, anchor="w")
        self.tree.column("status", width=70, anchor="center")
        self.tree.column("file_count", width=70, anchor="center")
        self.tree.column("output_file", width=260, anchor="w")

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="top", fill="both", expand=True, padx=12, pady=(4, 0))
        vsb.pack(side="right", fill="y", padx=(0, 12), pady=(4, 0))
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # ---- 详情面板 ----
        detail_frame = ttk.LabelFrame(self, text="详情", padding=8)
        detail_frame.pack(fill="x", padx=12, pady=8)
        self.detail_text = tk.Text(
            detail_frame, height=8, font=("Consolas", 9),
            state="disabled", relief="flat", borderwidth=0, wrap="word",
        )
        self.detail_text.pack(fill="both", expand=True)

        # ---- 按钮 ----
        btn_frame = ttk.Frame(self, padding=(12, 0, 12, 12))
        btn_frame.pack(fill="x")

        self.open_file_btn = ttk.Button(
            btn_frame, text="打开输出文件", command=self._on_open_file, state="disabled",
        )
        self.open_file_btn.pack(side="left", padx=(0, 4))
        self.open_folder_btn = ttk.Button(
            btn_frame, text="打开所在文件夹", command=self._on_open_folder, state="disabled",
        )
        self.open_folder_btn.pack(side="left", padx=4)
        ttk.Button(btn_frame, text="删除记录", command=self._on_delete).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="清空记录", command=self._on_clear).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="关闭", command=self.destroy).pack(side="right", padx=(4, 0))

    # ---- 数据 ----

    def _refresh_list(self):
        self.tree.delete(*self.tree.get_children())
        self._records, warning = task_history.read_history()
        self._history_intro.configure(text=warning or "本地任务记录（最多500条，最新在上方；预览不等于已导出）")
        # 最新任务在上方
        for r in reversed(self._records):
            status = task_history.status_label(r)
            self.tree.insert(
                "", "end", iid=r.get("id"),
                values=(
                    r.get("time", ""),
                    task_history.type_label(r),
                    status,
                    r.get("file_count", 0),
                    r.get("output_file", ""),
                ),
            )
        self._clear_detail()
        self._update_buttons()

    def _find_record(self, record_id):
        for r in self._records:
            if r.get("id") == record_id:
                return r
        return None

    def _selected_record(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self._find_record(sel[0])

    def _get_output_path(self):
        rec = self._selected_record()
        if not rec:
            return ""
        return task_history.output_path(rec)

    # ---- 展示 ----

    def _on_select(self, _event=None):
        rec = self._selected_record()
        if rec:
            self._show_detail(rec)
        else:
            self._clear_detail()
        self._update_buttons()

    def _show_detail(self, rec):
        lines = [
            f"时间：{rec.get('time', '')}",
            f"任务类型：{task_history.type_label(rec)}",
            f"状态：{task_history.status_label(rec)}",
            f"有效明细：{rec.get('valid_row_count', '未记录')}；结果数量：{rec.get('result_count', '未记录')}",
            f"异常：{rec.get('issue_count', '未记录')}；日期排除：{rec.get('excluded_count', '未记录')}",
            f"输入文件夹：{rec.get('input_folder', '')}",
            f"输入文件：{', '.join(rec.get('input_files', []))}",
            f"文件数：{rec.get('file_count', 0)}",
            f"输出文件：{rec.get('output_file', '')}",
            f"输出目录：{rec.get('output_folder', '')}",
            f"字段：{', '.join(rec.get('fields', []))}",
        ]
        if rec.get("error"):
            lines.append(f"错误：{rec.get('error')}")
        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", "\n".join(lines))
        self.detail_text.config(state="disabled")

    def _clear_detail(self):
        self.detail_text.config(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.config(state="disabled")

    def _update_buttons(self):
        rec = self._selected_record()
        path = self._get_output_path()
        has_file = bool(rec and os.path.isfile(path))
        folder = os.path.dirname(path) if path else ""
        has_folder = bool(rec and folder and os.path.isdir(folder))
        self.open_file_btn.config(state="normal" if has_file else "disabled")
        self.open_folder_btn.config(state="normal" if has_folder else "disabled")

    # ---- 操作 ----

    def _on_open_file(self):
        path = self._get_output_path()
        if path and os.path.isfile(path):
            os.startfile(path)
        else:
            messagebox.showinfo("提示", "输出文件不存在或已被移动。", parent=self)

    def _on_open_folder(self):
        rec = self._selected_record()
        path = task_history.output_path(rec) if rec else ""
        folder = os.path.dirname(path) if path else ""
        if folder and os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showinfo("提示", "输出文件夹不存在。", parent=self)

    def _on_delete(self):
        rec = self._selected_record()
        if not rec:
            messagebox.showinfo("提示", "请先选择一条记录。", parent=self)
            return
        if not messagebox.askyesno(
            "确认删除",
            "确定删除这条任务记录吗？\n（不会删除实际的 Excel 文件）",
            parent=self,
        ):
            return
        if not delete_record(rec.get("id")):
            messagebox.showwarning("记录未删除", "记录文件无法安全更新，原文件已保留。", parent=self)
        self._refresh_list()

    def _on_clear(self):
        if not self._records:
            messagebox.showinfo("提示", "没有可清空的记录。", parent=self)
            return
        if not messagebox.askyesno(
            "确认清空",
            "确定清空所有任务记录吗？\n（不会删除任何实际的 Excel 文件）",
            parent=self,
        ):
            return
        if not task_history.clear_records_safe():
            messagebox.showwarning("记录未清空", "记录文件无法安全更新，原文件已保留。", parent=self)
        self._refresh_list()


# ============================================================
#  v1.5.2: 关于软件窗口
# ============================================================

class AboutDialog(tk.Toplevel):
    """「关于软件」对话框 — 展示版本、运行方式与数据目录。"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("关于软件")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        data_dir = get_data_dir()

        main = ttk.Frame(self, padding=24)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text=APP_NAME, font=("Microsoft YaHei", 16, "bold")).pack()
        ttk.Label(
            main, text=APP_VERSION, font=("Microsoft YaHei", 10), foreground="#888",
        ).pack(pady=(2, 16))

        info = ttk.Frame(main)
        info.pack(fill="x")
        ttk.Label(info, text="运行方式：离线本地运行，无网络需求").pack(anchor="w")
        ttk.Label(info, text=f"数据目录：{data_dir}", wraplength=420).pack(anchor="w")

        ttk.Button(
            main, text="打开数据目录", command=lambda: os.startfile(data_dir),
        ).pack(pady=(20, 6))
        ttk.Button(main, text="关闭", command=self.destroy).pack(pady=(0, 4))

        self.update_idletasks()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")


# ============================================================
#  入口
# ============================================================


def main():
    from .display import enable_high_dpi
    enable_high_dpi()
    logs_dir = get_logs_dir()
    log_path = os.path.join(logs_dir, "startup.log")

    with open(log_path, "a", encoding="utf-8") as f:
        f.write("start\n")

    # ---- 主程序 Tk 生命周期 ----
    root = tk.Tk()

    with open(log_path, "a", encoding="utf-8") as f:
        f.write("create SummaryApp begin\n")
    from .desktop_ui import UnifiedApp
    app = UnifiedApp(root)
    with open(log_path, "a", encoding="utf-8") as f:
        f.write("create SummaryApp end\n")

    with open(log_path, "a", encoding="utf-8") as f:
        f.write("mainloop begin\n")
    root.mainloop()


if __name__ == "__main__":
    main()
