"""Pure preview rows and minimal history metadata, without Tk dependencies."""
from dataclasses import dataclass

from .analysis import date_trace_text
from .monthly import MonthlyResult
from .reconciliation import CATEGORIES


@dataclass(frozen=True)
class Preview:
    summary: list
    sources: list
    issues: list
    message: str


def build_preview(result):
    categories, summary = {}, []
    if isinstance(result, MonthlyResult):
        reports = [("本月", result.report)]
        for row in result.rows:
            summary.append((row.key, "月度汇总", str(row.amount), "", "", len(row.records)))
            for record in row.records:
                categories["本月", record.source] = "月度汇总"
        message = f"预览完成：{len(result.rows)} 个分组，{len(result.report.records)} 条有效明细"
        if result.report.excluded_rows:
            message += f"；外月跳过 {result.report.excluded_rows} 行（见来源明细「日期排除」）"
    else:
        reports = [("左侧", result.left), ("右侧", result.right)]
        for row in result.rows:
            summary.append((row.key, CATEGORIES[row.category],
                            "" if row.left_amount is None else str(row.left_amount),
                            "" if row.right_amount is None else str(row.right_amount),
                            "" if row.difference is None else str(row.difference),
                            len(row.left) + len(row.right)))
            for side, records in (("左侧", row.left), ("右侧", row.right)):
                for record in records:
                    categories[side, record.source] = CATEGORIES[row.category]
        message = "预览完成：" + "，".join(f"{CATEGORIES[c]} {n}" for c, n in result.counts.items())
        if result.all_consistent:
            message += "；全部一致（按当前容差）"
    sources, issues = [], []
    for side, report in reports:
        for record in report.records:
            source = record.source
            sources.append((side, categories[side, source], source.file, source.sheet,
                            source.row, record.key, str(record.amount),
                            date_trace_text(getattr(record, "raw_date", ""),
                                            getattr(record, "normalized_date", ""))))
        for excluded in getattr(report, "excluded", ()):
            source = excluded.source
            sources.append((side, "日期排除", source.file, source.sheet, source.row,
                            getattr(excluded, "raw_key", ""), "未校验",
                            date_trace_text(excluded.raw_date, excluded.normalized)))
        for issue in report.issues:
            source = issue.source
            issues.append((side, source.file, source.sheet, source.row, issue.code,
                           issue.message, issue.raw_key, issue.raw_amount,
                           getattr(issue, "raw_date", ""), getattr(issue, "normalized_date", "")))
    if not result.complete:
        ambiguous = 0 if isinstance(result, MonthlyResult) else result.counts["ambiguous"]
        message += f"；结果不完整：异常 {len(issues)} 项，重复歧义 {ambiguous} 个键，请核查"
    return Preview(summary, sources, issues, message)


def history_counts(result):
    reports = [result.report] if isinstance(result, MonthlyResult) else [result.left, result.right]
    return {"file_count": sum(len(report.input_files) for report in reports),
            "valid_row_count": sum(len(report.records) for report in reports),
            "result_count": len(result.rows), "issue_count": sum(len(report.issues) for report in reports),
            "excluded_count": sum(report.excluded_rows for report in reports), "complete": result.complete}
