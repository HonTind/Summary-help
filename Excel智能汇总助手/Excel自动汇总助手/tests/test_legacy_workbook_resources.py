"""Every legacy input workbook is released on success, skip and error paths."""
from contextlib import ExitStack, redirect_stdout
import io
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from Excel自动汇总助手 import excel_reader as reader, field_mapper as mapper, gui


class WorkbookResourceTests(unittest.TestCase):
    def test_aggregate_releases_on_success_skip_and_processing_error(self):
        for extension in (".xlsx", ".xlsm", ".xls"):
            for outcome in ("success", "skip", "error"):
                with self.subTest(extension=extension, outcome=outcome), tempfile.TemporaryDirectory() as folder, ExitStack() as stack:
                    book = MagicMock()
                    sheet = book.sheet_by_index.return_value if extension == ".xls" else book.active
                    sheet.nrows = 2
                    sheet.row_values.return_value = ["A", 2]
                    sheet.iter_rows.return_value = [("A", 2)]
                    closer = book.release_resources if extension == ".xls" else book.close
                    opener = stack.enter_context(patch.object(reader.xlrd, "open_workbook", return_value=book)
                                                 if extension == ".xls" else patch.object(reader, "load_workbook", return_value=book))
                    stack.enter_context(patch.object(reader, "_get_excel_files", return_value=["sample" + extension]))
                    stack.enter_context(patch.object(reader, "_get_effective_config", return_value=[{"name": "客户"}, {"name": "金额"}]))
                    detections = [SimpleNamespace(field_name="客户", col_index=0), SimpleNamespace(field_name="金额", col_index=1)]
                    stack.enter_context(patch.object(reader, "_detect_columns_xls" if extension == ".xls" else "detect_columns",
                                                    return_value=None if outcome == "skip" else detections))
                    stack.enter_context(patch.object(reader, "diagnose_failure", return_value=(["客户"], 0, None, False)))
                    stack.enter_context(patch.object(reader, "missing_required_fields", return_value=["金额"]))
                    stack.enter_context(patch.object(reader, "_print_field_error"))
                    stack.enter_context(redirect_stdout(io.StringIO()))
                    if outcome == "error":
                        stack.enter_context(patch.object(reader, "_add_row_to_total", side_effect=ValueError("synthetic")))
                        with self.assertRaisesRegex(ValueError, "synthetic"):
                            reader.aggregate_with_report(folder, folder)
                    else:
                        report = reader.aggregate_with_report(folder, folder)
                        self.assertEqual(report["processed_count"], 1 if outcome == "success" else 0)
                        self.assertEqual(len(report["skipped"]), 1 if outcome == "skip" else 0)
                    opener.assert_called_once()
                    closer.assert_called_once()

    def test_header_reader_releases_success_empty_and_error(self):
        for extension in (".xlsx", ".xls"):
            for outcome in ("success", "empty", "error"):
                with self.subTest(extension=extension, outcome=outcome), ExitStack() as stack:
                    book = MagicMock()
                    if extension == ".xls":
                        stack.enter_context(patch("xlrd.open_workbook", return_value=book))
                        book.sheet_by_index.return_value.nrows = 0 if outcome == "empty" else 1
                        book.sheet_by_index.return_value.row_values.return_value = ["客户", "金额"]
                        if outcome == "error":
                            book.sheet_by_index.side_effect = ValueError("synthetic")
                        closer = book.release_resources
                    else:
                        stack.enter_context(patch("openpyxl.load_workbook", return_value=book))
                        stack.enter_context(patch.object(mapper, "get_headers_for_display",
                                                       return_value=[] if outcome == "empty" else [(0, "客户")],
                                                       side_effect=ValueError("synthetic") if outcome == "error" else None))
                        closer = book.close
                    ok, _ = mapper.get_file_headers("sample" + extension)
                    self.assertEqual(ok, outcome == "success")
                    closer.assert_called_once()

    def test_gui_scan_releases_book_when_detection_raises(self):
        for extension in (".xlsx", ".xls"):
            with self.subTest(extension=extension), tempfile.TemporaryDirectory() as folder, ExitStack() as stack:
                book = MagicMock()
                stack.enter_context(patch("xlrd.open_workbook" if extension == ".xls" else "openpyxl.load_workbook", return_value=book))
                stack.enter_context(patch.object(gui, "_detect_columns_xls" if extension == ".xls" else "detect_columns",
                                                side_effect=ValueError("synthetic")))
                stack.enter_context(patch.object(gui, "get_logs_dir", return_value=folder))
                host = MagicMock()
                host._preview_ok, host._preview_warn = [], []
                host.selected_excel_files = {"sample" + extension}
                host.selected_folder.get.return_value = folder
                host.task_field_mappings = {}
                gui.SummaryApp._rescan_selected_files(host)
                self.assertEqual(len(host._preview_warn), 1)
                (book.release_resources if extension == ".xls" else book.close).assert_called_once()


if __name__ == "__main__":
    unittest.main()
