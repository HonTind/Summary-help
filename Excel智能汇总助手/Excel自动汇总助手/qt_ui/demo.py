"""Explicit in-memory UI demonstration, not a replacement business engine.

No file, configuration, history, network or license access. A future adapter
must call the existing engines; these synthetic rows never enter exports.
"""
from dataclasses import dataclass
from datetime import date
from decimal import Decimal
from typing import Protocol


@dataclass(frozen=True)
class PreviewRequest:
    mode: str = "monthly"
    files: tuple[str, ...] = ("October-A.xlsx", "October-B.xlsx")
    month: str = "2026-10"
    key: str = "客户"
    amount: str = "金额"
    date_filter: bool = False
    date_field: str = "日期"
    compare_mode: str = "客户总额"
    extra_rows: int = 0


@dataclass(frozen=True)
class TableData:
    headers: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]
    footer: tuple[str, ...] | None = None  # Presentation only, never a data row.


@dataclass(frozen=True)
class PreviewData:
    summary: TableData
    sources: TableData
    issues: TableData
    message: str


class PreviewBackend(Protocol):
    def preview(self, request: PreviewRequest) -> PreviewData: ...


class DemoBackend:
    """Interactive sample selection/filtering with clearly labelled results."""
    samples = ("October-A.xlsx", "October-B.xlsx")

    def preview(self, request):
        if not request.files:
            raise ValueError("请选择至少一份演示文件。")
        if request.extra_rows not in (0, 201):
            raise ValueError("只允许默认样例或201条额外演示数据。")
        if any(name not in self.samples for name in request.files):
            raise ValueError("本样板只允许预置演示文件，不读取真实账单。")
        if request.key not in ("客户", "单据号", "#1") or request.amount not in ("金额", "#2"):
            raise ValueError("演示字段仅包含客户、单据号、金额及示例列号。真实字段将由现有引擎识别。")
        if request.mode not in ("regular", "monthly", "compare"):
            raise ValueError("未知任务模式。")
        if request.mode == "compare" and set(request.files) != set(self.samples):
            raise ValueError("对账演示需同时选择两份内置样例，一份为左侧、一份为右侧。")
        if request.mode == "monthly":
            try:
                if len(request.month) != 7 or request.month[4] != "-":
                    raise ValueError
                date(int(request.month[:4]), int(request.month[5:]), 1)
            except ValueError as error:
                raise ValueError("月份请填写 YYYY-MM，例如 2026-10。") from error
            if request.date_filter and request.date_field not in ("日期", "#3"):
                raise ValueError("演示日期字段为“日期”或 #3。")
        lines = []
        for name in dict.fromkeys(request.files):
            rows = (("示例客户A", "1280.00"), ("示例客户B", "890.00")) if name == self.samples[0] else (("示例客户A", "0.50"),)
            if request.mode == "monthly" and request.date_filter and request.month != "2026-10":
                rows = ()
            for number, (key, amount) in enumerate(rows, 2):
                if request.key == "单据号":
                    key = "示例单据001" if key == "示例客户A" else "示例单据002"
                lines.append((name, str(number), key, amount, "2026-10-01"))
        if self.samples[0] in request.files and not (request.mode == "monthly" and request.date_filter and request.month != "2026-10"):
            for index in range(request.extra_rows):
                key = f"示例分页客户{index + 1:03d}"
                lines.append((self.samples[0], str(index + 10), key, "1.00", "2026-10-01"))
        totals, counts = {}, {}
        for _name, _row, key, amount, _date in lines:
            totals[key] = totals.get(key, Decimal(0)) + Decimal(amount)
            counts[key] = counts.get(key, 0) + 1
        total = sum(totals.values(), Decimal(0))
        if request.mode == "compare":
            keys = ("示例单据001", "示例单据002") if request.compare_mode == "单据逐笔" else ("示例客户A", "示例客户B")
            summary = TableData(("客户 / 单据", "结果", "左金额", "右金额", "差额"), (
                (keys[0], "一致", "1280.50", "1280.50", "0.00"),
                (keys[1], "金额差异", "890.00", "885.00", "5.00")))
            message = f"演示数据 · {request.compare_mode}。真实匹配与导出尚未接入。"
            sources = TableData(("侧别", "文件", "原行号", "客户 / 单据", "金额"), (
                ("左", self.samples[0], "2", keys[0], "1280.50"),
                ("左", self.samples[0], "3", keys[1], "890.00"),
                ("右", self.samples[1], "2", keys[0], "1280.50"),
                ("右", self.samples[1], "3", keys[1], "885.00")))
            extra_keys = tuple(f"示例分页{'单据' if request.compare_mode == '单据逐笔' else '客户'}{index + 1:03d}"
                               for index in range(request.extra_rows))
            summary = TableData(summary.headers, summary.rows + tuple((key, "一致", "1.00", "1.00", "0.00") for key in extra_keys))
            sources = TableData(sources.headers, sources.rows + tuple(
                (side, name, str(index + 10), key, "1.00")
                for index, key in enumerate(extra_keys)
                for side, name in (("左", self.samples[0]), ("右", self.samples[1]))))
        else:
            summary = TableData((request.key, "金额（精确文本）", "来源行数"),
                tuple((key, f"{amount:.2f}", str(counts[key])) for key, amount in totals.items()),
                ("合计", f"{total:.2f}", str(len(lines))))
            message = "演示数据，仅用于展示界面。"
            sources = TableData(("文件", "原行号", request.key, "金额", "日期"), tuple(lines))
        if request.extra_rows:
            message += "已启用内存大样例，可体验分页。"
        return PreviewData(summary, sources,
            TableData(("文件", "原行号", "异常", "说明"), ()), message)
