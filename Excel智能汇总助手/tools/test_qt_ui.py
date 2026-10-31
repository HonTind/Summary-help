"""Headless acceptance for the optional Qt sample; no production data access."""
import argparse
import importlib
import json
import os
from pathlib import Path
import socket
import subprocess
import sys
import time
import unittest
import urllib.request
import webbrowser
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
PROJECT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT))
from PySide6.QtCore import Qt, QEvent
from PySide6.QtWidgets import QApplication, QMessageBox
from Excel自动汇总助手.qt_ui.demo import DemoBackend, PreviewRequest, TableData
from Excel自动汇总助手.qt_ui.theme import configure_application
from Excel自动汇总助手.qt_ui.window import PreviewWindow


class DemoTests(unittest.TestCase):
    def test_exact_sample_total_and_separate_presentation_footer(self):
        data = DemoBackend().preview(PreviewRequest())
        self.assertEqual(data.summary.rows[0], ("示例客户A", "1280.50", "2"))
        self.assertEqual(data.summary.footer, ("合计", "2170.50", "3"))
        self.assertEqual(len(data.summary.rows), 2)

    def test_month_label_does_not_filter_without_explicit_toggle(self):
        backend = DemoBackend()
        self.assertEqual(backend.preview(PreviewRequest(month="2026-09")).summary.footer[1], "2170.50")
        self.assertEqual(backend.preview(PreviewRequest(month="2026-09", date_filter=True)).summary.rows, ())

    def test_real_file_names_invalid_fields_and_dates_are_rejected(self):
        for request in (PreviewRequest(files=("real.xlsx",)), PreviewRequest(files=()),
                        PreviewRequest(month="2026-13"), PreviewRequest(key="未知"),
                        PreviewRequest(date_filter=True, date_field="未知")):
            with self.subTest(request=request), self.assertRaises(ValueError):
                DemoBackend().preview(request)

    def test_both_compare_modes_are_explicit_sample_fixtures(self):
        for mode, key in (("客户总额", "示例客户A"), ("单据逐笔", "示例单据001")):
            data = DemoBackend().preview(PreviewRequest(mode="compare", compare_mode=mode))
            self.assertEqual(data.summary.rows[0][0], key)
            self.assertEqual(data.summary.rows[1][-1], "5.00")
            self.assertIn("尚未接入", data.message)
            self.assertEqual(data.sources.rows[0][3], key)
            self.assertEqual(data.sources.rows[0][-1], data.summary.rows[0][2])
            self.assertEqual(data.sources.rows[2][-1], data.summary.rows[0][3])

    def test_large_sample_is_bounded_and_provenance_stays_consistent(self):
        data = DemoBackend().preview(PreviewRequest(extra_rows=201))
        self.assertEqual(len(data.summary.rows), 203)
        self.assertEqual(data.summary.footer[1], "2371.50")
        self.assertEqual(len(data.sources.rows), 204)


class WidgetTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])
        configure_application(cls.app)

    def setUp(self):
        def no_network(*_args, **_kwargs):
            raise AssertionError("Qt sample tried to access network/external browser")
        for target, name in ((socket.socket, "connect"), (socket, "create_connection"),
                             (urllib.request, "urlopen"), (webbrowser, "open")):
            mock = patch.object(target, name, side_effect=no_network)
            mock.start()
            self.addCleanup(mock.stop)
        self.window = PreviewWindow()
        self.window.resize(1168, 824)
        self.window.show()
        self.app.processEvents()
        self.addCleanup(self.cleanup)

    def cleanup(self):
        self.window.close()
        self.app.processEvents()

    def drain(self):
        deadline = time.monotonic() + 2
        while self.window.busy and time.monotonic() < deadline:
            self.app.processEvents()
            time.sleep(.005)
        self.app.processEvents()
        self.assertFalse(self.window.busy)

    def test_all_pages_and_modes_route_without_production_imports(self):
        for page in self.window.PAGES:
            self.window.show_page(page)
            self.assertEqual(self.window.pages.currentWidget(), self.window.page_widgets[page])
        for mode in self.window.MODES:
            self.window.select_mode(mode)
            self.assertEqual(self.window.mode, mode)
            self.assertFalse(self.window.export_button.isEnabled())
        for module in ("gui", "desktop_ui", "utils", "history", "template_store"):
            self.assertNotIn("Excel自动汇总助手." + module, sys.modules)
        self.assertNotIn("PySide6.QtNetwork", sys.modules)
        self.assertNotIn("PySide6.QtWebEngineWidgets", sys.modules)

    def test_preview_freezes_controls_invalidates_and_unlocks(self):
        self.window.key.setCurrentText("客户")
        self.window.month.setCurrentText("2026-09")
        self.assertIsNone(self.window.result)
        self.window.preview_button.click()
        self.assertTrue(self.window.busy)
        self.assertFalse(self.window.month.isEnabled())
        self.assertFalse(self.window.nav_buttons["history"].isEnabled())
        self.window.select_mode("compare")
        self.assertEqual(self.window.mode, "monthly")
        self.drain()
        self.assertEqual(self.window.result.summary.footer[1], "2170.50")
        self.assertTrue(self.window.export_button.isEnabled())
        self.window.date_check.setChecked(True)
        self.assertFalse(self.window.export_button.isEnabled())
        self.assertTrue(self.window.date_field.isVisible())

    def test_error_state_has_no_old_exportable_result(self):
        self.window.month.setCurrentText("invalid")
        self.window.preview_button.click()
        self.drain()
        self.assertIsNone(self.window.result)
        self.assertFalse(self.window.export_button.isEnabled())
        self.assertIn("YYYY-MM", self.window.info.text())

    def test_template_memory_and_help_search_jump(self):
        self.window.template.setCurrentText("演示模板2")
        self.window.key.setCurrentText("#1")
        self.window.save_memory_template()
        self.assertEqual(self.window.memory_templates["演示模板2"], ("#1", "金额"))
        self.window.key.setCurrentText("客户")
        self.window.load_memory_template()
        self.assertEqual(self.window.key.currentText(), "#1")
        self.window.show_page("help")
        self.window.help_search.setText("单据")
        self.assertGreater(self.window.help_topics.count(), 0)
        self.window.select_mode("monthly")
        self.window.context_help()
        self.assertEqual(self.window.help_search.text(), "")
        self.assertTrue(self.window.help_text.toPlainText().startswith("月度汇总"))
        self.window.jump_help()
        self.assertEqual(self.window.mode, "monthly")

    def test_export_explains_boundary_without_any_writer(self):
        with patch.object(QMessageBox, "information") as dialog:
            self.window.export_button.click()
        self.assertIn("不会生成", dialog.call_args.args[2])
        self.assertIn("尚未接入", dialog.call_args.args[1])

    def test_paging_does_not_count_display_total_as_data(self):
        panel = self.window.table_panels[0]
        panel.set_data(TableData(("客户", "金额", "来源行数"),
            tuple((f"示例{i}", "1.00", "1") for i in range(101)), ("合计", "101.00", "101")))
        self.assertEqual(panel.table.rowCount(), 100)
        self.assertTrue(panel.next.isEnabled())
        panel.next.click()
        self.assertEqual(panel.table.rowCount(), 2)
        self.assertEqual(panel.table.item(1, 0).data(Qt.ItemDataRole.UserRole), "presentation-footer")
        self.assertFalse(panel.next.isEnabled())
        self.assertEqual(len(panel.data.rows), 101)

    def test_main_actions_and_pager_fit_narrow_window(self):
        self.window.resize(840, 580)
        self.app.processEvents()
        for widget in (self.window.preview_button, self.window.export_button, self.window.help_button,
                       self.window.table_panels[0].count, self.window.table_panels[0].next):
            position = widget.mapTo(self.window, widget.rect().topLeft())
            self.assertGreaterEqual(position.x(), 0)
            self.assertGreaterEqual(position.y(), 0)
            self.assertLessEqual(position.x() + widget.width(), self.window.width())
            self.assertLessEqual(position.y() + widget.height(), self.window.height())
        self.assertEqual(self.window.width(), 840)
        self.assertEqual(self.window.height(), 580)
        self.assertEqual(self.window.footer_hint.text(), "仅演示 · 无文件写入")
        self.assertLessEqual(self.window.footer_hint.fontMetrics().horizontalAdvance(self.window.footer_hint.text()),
                             self.window.footer_hint.width())

    def test_large_sample_is_reachable_and_timer_is_reused(self):
        timer = self.window.timer
        self.window.large_sample.setChecked(True)
        self.window.preview()
        self.drain()
        panel = self.window.table_panels[0]
        self.assertTrue(panel.next.isEnabled())
        panel.next.click()
        self.assertEqual(panel.page, 1)
        self.window.preview()
        self.drain()
        self.assertIs(self.window.timer, timer)

    def test_small_high_density_work_area_keeps_actions_inside(self):
        window = PreviewWindow(available_size=(960, 540))
        self.addCleanup(window.close)
        window.show()
        self.app.processEvents()
        self.assertLessEqual(window.height(), 480)
        for button in (window.preview_button, window.export_button):
            position = button.mapTo(window, button.rect().topLeft())
            self.assertLessEqual(position.y() + button.height(), window.height())

    def test_replacing_files_drops_dead_icon_labels_before_dpi_refresh(self):
        before = len(self.window._image_labels)
        self.window.selected_files = [DemoBackend.samples[0]]
        self.window.refresh_files()
        self.app.sendPostedEvents(None, QEvent.Type.DeferredDelete)
        self.assertEqual(len(self.window._image_labels), before - 1)
        self.app.sendEvent(self.window, QEvent(QEvent.Type.DevicePixelRatioChange))
        self.assertTrue(all(not label.pixmap().isNull() for label, _icon, _size in self.window._image_labels))


def probe():
    app = QApplication([])
    configure_application(app)
    window = PreviewWindow()
    window.resize(1168, 824)
    window.show()
    app.processEvents()
    print(json.dumps({"dpr": window.devicePixelRatioF(), "logical_client": [window.width(), window.height()],
                      "icon_valid": not window.icons.get("workbench").isNull()}), flush=True)
    window.close()
    app.processEvents()


if __name__ == "__main__":
    if "--probe" in sys.argv:
        probe()
    else:
        unittest.main()
