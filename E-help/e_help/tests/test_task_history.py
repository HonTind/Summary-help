"""Local synthetic metadata only; never read the user's history or business files."""
import importlib
import json
from pathlib import Path
import sys
import tempfile
import threading
import unittest
from unittest.mock import patch

PROJECT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(PROJECT))
history = importlib.import_module("e_help.history")


class HistoryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="ehelper-history-test-")
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.path = self.folder / "history.json"
        self.mock = patch.object(history, "get_history_dir", return_value=str(self.folder))
        self.mock.start()
        self.addCleanup(self.mock.stop)

    def seed(self, data):
        self.path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        return self.path.read_bytes()

    def test_old_version_one_and_missing_version_are_not_rewritten(self):
        old = {"id": "legacy", "status": "success", "time": "2020-01-01", "input_files": ["synthetic.xlsx"],
               "fields": ["示例字段"], "file_count": 1, "error": "old synthetic note"}
        for data in ({"version": 1, "records": [old]}, {"records": [old]}):
            before = self.seed(data)
            self.assertEqual(history.load_history(), [old])
            self.assertEqual(self.path.read_bytes(), before)
            self.assertEqual(history.type_label(old), "常规汇总")

    def test_append_preserves_old_record_and_unknown_root_metadata(self):
        old = {"id": "old", "fields": ["legacy"], "status": "failed"}
        self.seed({"version": 1, "records": [old], "custom": {"legacy": True}})
        self.assertTrue(history.save_task("monthly", "previewed", phase="preview", result_count=3))
        data = json.loads(self.path.read_text(encoding="utf-8"))
        self.assertEqual(data["records"][0], old)
        self.assertEqual(data["custom"], {"legacy": True})
        self.assertEqual(data["version"], 1)

    def test_new_records_whitelist_metadata_and_discard_preview_path(self):
        marker = "SYNTHETIC_PRIVATE_CONTENT"
        self.assertTrue(history.save_task("monthly", "previewed", phase="preview", file_count=2,
                                          output_path=str(self.folder / "never-generated.xlsx"), result_count=3,
                                          valid_row_count=4, issue_count=0, excluded_count=1, complete=True,
                                          input_files=[marker], error=marker, amount=marker, fields=[marker]))
        record = history.load_history()[0]
        self.assertNotIn(marker, self.path.read_text(encoding="utf-8"))
        self.assertEqual(record["output_file"], "")
        self.assertEqual(record["result_count"], 3)
        self.assertEqual(set(record), {"id", "time", "task_type", "status", "phase", "file_count", "result_count",
                                      "valid_row_count", "issue_count", "excluded_count", "complete",
                                      "output_file", "output_folder"})

    def test_broken_unknown_or_invalid_history_never_overwritten(self):
        cases = [b"{broken", b'[]', b'{"version":2,"records":[]}', b'{"version":true,"records":[]}',
                 b'{"version":1,"records":[{}]}', b'{"version":1,"records":[{"id":12}]}',
                 b'{"records":[{"id":"same"},{"id":"same"}]}']
        for content in cases:
            self.path.write_bytes(content)
            self.assertTrue(history.read_history()[1])
            self.assertFalse(history.save_task("regular", "failed", phase="summary"))
            history.add_record({"status": "success"})
            self.assertFalse(history.delete_record("same"))
            self.assertFalse(history.clear_records_safe())
            history.clear_records()
            self.assertEqual(self.path.read_bytes(), content)

    def test_failed_replace_or_serialization_preserves_original_and_removes_owned_temp(self):
        before = self.seed({"version": 1, "records": [{"id": "old"}]})
        with patch.object(history.os, "replace", side_effect=PermissionError("synthetic denied")):
            self.assertFalse(history.save_task("monthly", "cancelled", phase="preview"))
        self.assertEqual(self.path.read_bytes(), before)
        self.assertEqual(list(self.folder.glob(".history-*.tmp")), [])
        self.assertFalse(history.append_record({"id": "bad", "invalid": float("nan")}))
        self.assertEqual(self.path.read_bytes(), before)
        self.assertEqual(list(self.folder.glob(".history-*.tmp")), [])

    def test_read_permission_failure_does_not_raise_or_start_write(self):
        self.seed({"records": []})
        with patch("builtins.open", side_effect=PermissionError("private path not echoed")):
            self.assertEqual(history.load_history(), [])
            self.assertTrue(history.read_history()[1])
            self.assertFalse(history.save_task("regular", "failed", phase="summary"))

    def test_legacy_add_return_and_id_completion_delete_clear(self):
        record = {"status": "success"}
        self.assertIsNone(history.add_record(record))
        self.assertTrue(record["id"])
        self.assertTrue(record["time"])
        self.assertFalse(history.append_record(record))
        self.assertTrue(history.delete_record(record["id"]))
        self.assertFalse(history.delete_record(record["id"]))
        history.add_record({"id": "next"})
        self.assertIsNone(history.clear_records())
        self.assertEqual(history.load_history(), [])

    def test_retention_preserves_existing_five_hundred_limit(self):
        self.seed({"version": 1, "records": [{"id": str(i)} for i in range(history.MAX_RECORDS)]})
        self.assertTrue(history.save_task("document", "failed", phase="preview"))
        records = history.load_history()
        self.assertEqual(len(records), 500)
        self.assertEqual(records[0]["id"], "1")

    def test_type_filters_and_readonly_labels_for_both_modes(self):
        records = [{"id": "old"}]
        for kind in history.TASK_TYPES:
            records.append({"id": kind, "task_type": kind, "status": "failed", "phase": "preview"})
        self.assertEqual(len(history.filter_records(records)), 5)
        self.assertEqual(len(history.filter_records(records, "regular")), 2)
        for kind in ("monthly", "customer", "document"):
            self.assertEqual(history.filter_records(records, kind)[0]["id"], kind)
        self.assertEqual(history.status_label({"status": "previewed"}), "预览完成")
        self.assertIn("不完整", history.status_label({"status": "success", "complete": False}))
        self.assertIn("导出", history.status_label({"status": "cancelled", "phase": "export"}))
        self.assertEqual(history.type_label({"task_type": []}), "其他任务")
        self.assertEqual(history.status_label({"status": []}), "未知状态")

    def test_safe_output_path_never_returns_url_unc_script_or_preview(self):
        local = {"status": "success", "output_folder": str(self.folder), "output_file": "saved.xlsx"}
        self.assertEqual(history.output_path(local), str(self.folder / "saved.xlsx"))
        for changes in ({"status": "previewed"}, {"output_file": "script.exe"},
                        {"output_folder": "https://example.invalid/"},
                        {"output_folder": "\\\\server\\share"}, {"output_folder": None},
                        {"output_file": "bad\x00.xlsx"}):
            self.assertEqual(history.output_path(dict(local, **changes)), "")

    def test_invalid_metadata_types_and_phase_pairs_rejected(self):
        for kind, status, phase in (([], "failed", "summary"), ("unknown", "failed", "preview"),
                                    ("regular", "previewed", "preview"), ("monthly", "success", "preview"),
                                    ("customer", "previewed", "export")):
            self.assertFalse(history.save_task(kind, status, phase=phase))
        self.assertTrue(history.save_task("customer", "failed", phase="preview", file_count=True,
                                          result_count=-1, valid_row_count="10"))
        record = history.load_history()[0]
        self.assertEqual(record["file_count"], 0)
        self.assertEqual(record["result_count"], 0)
        self.assertEqual(record["valid_row_count"], 0)

    def test_process_local_concurrent_appends_keep_all_records(self):
        results = []
        threads = [threading.Thread(target=lambda: results.append(
            history.save_task("regular", "failed", phase="summary"))) for _ in range(12)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join(3)
        self.assertTrue(all(not thread.is_alive() for thread in threads))
        self.assertEqual(results, [True] * 12)
        self.assertEqual(len({record["id"] for record in history.load_history()}), 12)


if __name__ == "__main__":
    unittest.main()
