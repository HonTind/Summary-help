"""Behavior contracts at the extracted presentation and worker boundaries."""
from copy import deepcopy
from decimal import Decimal
from pathlib import Path
import subprocess
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from Excel自动汇总助手.analysis import Source, Record, ReadReport, Issue, ExcludedRow
from Excel自动汇总助手.monthly import Aggregate, MonthlyResult
from Excel自动汇总助手.reconciliation import compare_reports
from Excel自动汇总助手.workbench_view import build_preview, history_counts


class PresentationTests(unittest.TestCase):
    def test_monthly_preview_retains_date_trace_issues_and_minimal_counts(self):
        source = Source("synthetic.xlsx", "Sheet1", 2)
        row = Record("A", Decimal("0.30"), source, ("A",), "2026/10/01", "2026-10-01")
        report = ReadReport(records=[row], files_loaded=1, input_files=["synthetic.xlsx"],
                            issues=[Issue(source, "invalid_date", "日期无效", "B", "1", "bad")],
                            excluded_rows=1,
                            excluded=[ExcludedRow(Source("synthetic.xlsx", "Sheet1", 4),
                                                  "2026/09/01", "2026-09-01", "C", "bad")])
        result = MonthlyResult("2026-10", {}, report, [Aggregate("A", row.amount, (row,))])
        before = deepcopy(result)
        preview = build_preview(result)
        self.assertEqual(preview.summary, [("A", "月度汇总", "0.30", "", "", 1)])
        self.assertEqual(preview.sources[0][:7], ("本月", "月度汇总", "synthetic.xlsx", "Sheet1", 2, "A", "0.30"))
        self.assertIn("2026/10/01", preview.sources[0][7])
        self.assertIn("2026-10-01", preview.sources[0][7])
        self.assertEqual(preview.sources[1][1], "日期排除")
        self.assertEqual(preview.sources[1][6], "未校验")
        self.assertEqual(preview.issues[0][4:], ("invalid_date", "日期无效", "B", "1", "bad", ""))
        self.assertIn("外月跳过 1 行", preview.message)
        self.assertIn("结果不完整：异常 1 项", preview.message)
        self.assertEqual(history_counts(result), dict(file_count=1, valid_row_count=1, result_count=1,
                                                      issue_count=1, excluded_count=1, complete=False))
        self.assertEqual(result, before)

    def test_reconciliation_categories_zero_missing_and_ambiguity(self):
        def report(side, values):
            return ReadReport(records=[Record(key, Decimal(amount), Source(side, "S", i), (key,))
                                       for i, (key, amount) in enumerate(values, 2)],
                              files_loaded=1, input_files=[side])
        left = report("left.xlsx", [("A", "0"), ("B", "2"), ("C", "3"), ("D", "1"), ("D", "2")])
        right = report("right.xlsx", [("A", "0"), ("B", "1"), ("E", "4"), ("D", "3")])
        for mode in ("customer", "document"):
            with self.subTest(mode=mode):
                result = compare_reports(left, right, mode)
                before = deepcopy(result)
                preview = build_preview(result)
                rows = {row[0]: row for row in preview.summary}
                self.assertEqual(rows["A"][1:5], ("一致", "0", "0", "0"))
                self.assertEqual(rows["B"][1:5], ("金额差异", "2", "1", "1"))
                self.assertEqual(rows["C"][1:5], ("仅左", "3", "", ""))
                self.assertEqual(rows["E"][1:5], ("仅右", "", "4", ""))
                self.assertEqual(rows["D"][1], "重复歧义" if mode == "document" else "一致")
                self.assertEqual(len(preview.sources), 9)
                self.assertEqual(preview.issues, [])
                self.assertEqual(result, before)

    def test_non_ui_modules_import_without_loading_tkinter(self):
        project = str(Path(__file__).resolve().parents[2])
        code = ("import sys; from Excel自动汇总助手 import job_runner, template_store, workbench_view; "
                "assert 'tkinter' not in sys.modules")
        result = subprocess.run([sys.executable, "-B", "-c", code], cwd=project,
                                capture_output=True, text=True, timeout=15)
        self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == "__main__":
    unittest.main()
