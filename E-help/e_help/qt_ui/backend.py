"""Offline adapter to existing engines. No Qt/Tk, network or IO at import."""
from contextlib import redirect_stdout
from copy import deepcopy
from dataclasses import dataclass, field
import io
import json
import os
from pathlib import Path
import tempfile

from ..analysis import SheetRules, check_cancel
from .. import field_mapper, history
from ..excel_reader import aggregate_with_report
from ..excel_writer import generate_summary
from ..monthly import MonthlyResult, make_template, run_monthly
from ..processor import sort_by_amount
from ..reconciliation import run_reconciliation
from ..template_store import load_templates, save_template
from ..workbench_export import export_result
from ..workbench_view import build_preview, history_counts
from .demo import PreviewData, TableData


def money_text(value):
    """Display at least two decimals without rounding away any precision."""
    value = format(value, "f")
    whole, separator, fraction = value.partition(".")
    return whole + "." + fraction.ljust(2, "0")


@dataclass(frozen=True)
class Request:
    mode: str
    paths: tuple
    right_paths: tuple = ()
    rules: SheetRules = field(default_factory=SheetRules)
    right_rules: SheetRules = field(default_factory=SheetRules)
    month: str = "2026-10"
    template_name: str = "标准月度规则"
    output_prefix: str = "月度汇总"
    compare_mode: str = "customer"
    tolerance: str = "0"
    mappings: dict = field(default_factory=dict)


@dataclass
class Result:
    request: Request
    raw: object
    preview: PreviewData
    complete: bool
    counts: dict


class Backend:
    def __init__(self, base_dir, config_dir, logs_dir):
        self.base_dir, self.config_dir, self.logs_dir = map(Path, (base_dir, config_dir, logs_dir))
        self.template_path = self.config_dir / "monthly_templates.json"

    def initialize(self):
        self.preferences()
        path = self.config_dir / field_mapper.FIELD_CONFIG_FILENAME
        if path.exists():
            with path.open(encoding="utf-8") as handle:
                config = json.load(handle)
            entries = config if isinstance(config, list) else [{"name": name, "keywords": words} for name, words in config.items()] if isinstance(config, dict) else None
            if not entries or any(not isinstance(row, dict) or not isinstance(row.get("name"), str)
                                  or not row["name"].strip() or not isinstance(row.get("keywords"), list)
                                  or any(not isinstance(word, str) for word in row["keywords"]) for row in entries):
                raise ValueError("字段配置无效，原文件已保留。")
        field_mapper.load_settings(str(self.config_dir))
        field_mapper.load_field_config(str(self.config_dir))

    def preview(self, request, cancel=None, progress=lambda _text: None):
        check_cancel(cancel)
        if not request.paths:
            raise ValueError("请先选择输入文件。")
        if request.mode == "regular":
            if any(not Path(path).is_file() for path in request.paths):
                raise ValueError("选中的输入文件已移动或删除，请重新选择；已阻止导出。")
            parents = {Path(path).resolve().parent for path in request.paths}
            if len(parents) != 1:
                raise ValueError("常规汇总请选择同一文件夹中的文件。")
            progress("正在自动识别字段并读取文件…")
            self.logs_dir.mkdir(parents=True, exist_ok=True)
            # The legacy engine prints diagnostics; no worker touches Qt widgets.
            with redirect_stdout(io.StringIO()):
                report = aggregate_with_report(str(next(iter(parents))), str(self.logs_dir),
                    [Path(path).name for path in request.paths], deepcopy(request.mappings))
            check_cancel(cancel)
            if report["skipped"]:
                names = "、".join(row["file"] for row in report["skipped"])
                raise ValueError(f"以下文件未成功读取，已阻止导出：{names}。请检查字段映射或文件内容。")
            if not report["file_count"] or not report["total"]:
                raise ValueError("没有可汇总的数据。")
            names = report["field_names"]
            report["sorted"] = sort_by_amount(report["total"], names)
            rows = tuple(tuple(values.get(name, "") for name in names) for _, values in report["sorted"])
            amount = sum(values.get(names[1], 0) for values in report["total"].values())
            footer = ("合计", str(amount), *("" for _ in names[2:]))
            preview = PreviewData(TableData(tuple(names), rows, footer),
                TableData(("输入文件", "说明"), tuple((Path(p).name, "原常规引擎不提供逐行来源") for p in request.paths)),
                TableData(("异常",), ()), f"预览完成：{report['file_count']} 个文件，{len(rows)} 个分组。")
            return Result(request, report, preview, True,
                          {"file_count": report["file_count"], "result_count": len(rows), "complete": True})
        if request.mode == "monthly":
            template = make_template(request.template_name, request.rules, request.output_prefix)
            raw = run_monthly(request.paths, template, request.month, cancel=cancel,
                              progress=lambda n, total, name: progress(f"读取 {n}/{total}：{name}"))
        elif request.mode == "compare":
            if not request.right_paths:
                raise ValueError("请选择右侧账单。")
            raw = run_reconciliation(request.paths, request.right_paths, request.rules, request.right_rules,
                mode=request.compare_mode, tolerance=request.tolerance, cancel=cancel,
                progress=lambda side, n, total, name: progress(f"{side} {n}/{total}：{name}"))
        else:
            raise ValueError("未知任务模式。")
        view = build_preview(raw)
        if isinstance(raw, MonthlyResult):
            label = str(request.rules.key_column) if not request.rules.key_fields else "组合键"
            summary = TableData((label, "金额（精确文本）", "来源行数"),
                tuple((r.key, money_text(r.amount), len(r.records)) for r in raw.rows),
                ("合计", money_text(raw.total), len(raw.report.records)))
        else:
            summary = TableData(("键", "分类", "左金额", "右金额", "左减右", "来源行数"), tuple(view.summary))
        preview = PreviewData(summary,
            TableData(("侧", "分类", "文件", "工作表", "行", "键", "金额（精确文本）", "日期"), tuple(view.sources)),
            TableData(("侧", "文件", "工作表", "行", "代码", "说明", "原始键", "原始金额", "原始日期", "规范日期"), tuple(view.issues)), view.message)
        return Result(request, raw, preview, raw.complete, history_counts(raw))

    def export(self, result, directory, cancel=None):
        if result.request.mode != "regular":
            return Path(export_result(result.raw, directory, cancel=cancel))
        check_cancel(cancel)
        directory = Path(directory)
        if not directory.is_dir():
            raise ValueError("输出目录不存在。")
        # Exclusive creation and a file handle avoid a name-check/save race.
        for index in range(10000):
            path = directory / ("汇总.xlsx" if index == 0 else f"汇总({index}).xlsx")
            try:
                handle = path.open("xb")
                break
            except FileExistsError:
                continue
        else:
            raise ValueError("无法生成唯一输出文件名。")
        try:
            with handle:
                generate_summary(result.raw["sorted"], result.raw["total"], handle, result.raw["field_names"])
        except BaseException:
            path.unlink(missing_ok=True)  # Only the new output exclusively created here.
            raise
        return path

    def templates(self):
        return load_templates(self.template_path)

    def save_template(self, template, *, allow_update=False):
        return save_template(self.template_path, template, allow_update=allow_update)

    def preferences(self):
        path = self.config_dir / "settings.json"
        if not path.exists():
            return {}
        with path.open(encoding="utf-8") as handle:
            data = json.load(handle)
        if not isinstance(data, dict):
            raise ValueError("设置文件结构无效，原文件已保留。")
        if "field_mapping" in data and not isinstance(data["field_mapping"], dict):
            raise ValueError("设置中的字段映射无效，原文件已保留。")
        return data

    def save_preferences(self, mode, directory):
        data = self.preferences()  # Never replace corrupt/unknown data with {}.
        data.update(output_mode=mode, output_dir=directory)
        self.config_dir.mkdir(parents=True, exist_ok=True)
        temporary = None
        try:
            with tempfile.NamedTemporaryFile("w", encoding="utf-8", dir=self.config_dir,
                                              prefix=".qt-settings-", suffix=".tmp", delete=False) as handle:
                temporary = Path(handle.name)
                json.dump(data, handle, ensure_ascii=False, indent=2, allow_nan=False)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, self.config_dir / "settings.json")
            temporary = None
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)

    def record(self, request, status, phase, *, result=None, output=""):
        kind = "regular" if request.mode == "regular" else "monthly" if request.mode == "monthly" else request.compare_mode
        # Regular history retains its summary-phase contract even for the new preview UI.
        if kind == "regular":
            phase = "summary"
            if status in ("previewed", "incomplete"):
                return True
        counts = dict(result.counts) if result else {"file_count": len(request.paths) + len(request.right_paths)}
        return history.save_task(kind, status, phase=phase, output_path=str(output), **counts)
