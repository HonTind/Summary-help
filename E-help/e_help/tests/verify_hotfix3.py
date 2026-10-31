"""
v1.5.2 Hotfix3 — 文件选择机制重构 验证测试

  核心验证点：
    - read_and_aggregate 新增 file_names 参数，只汇总指定文件（T1-T4）
    - ExcelFileSelectDialog 复选框列表 / 全选 / 全不选 / 反选 / 计数器（T5-T11）
    - _on_folder_selected 改为「扫描 → 弹窗勾选 → 仅勾选文件参与汇总」（T12-T16）
    - _on_start_summary 只使用 self.selected_excel_files，不再次扫描整个文件夹（T17-T19）
    - 字段映射 / 字段规则 只对已勾选文件重新识别（T20-T21）
    - 版本 / 端到端 / 全选保持旧行为（T22-T24）

  运行:
    cd "Day07(小试牛刀)源文件"
    python -m e_help.tests.verify_hotfix3
"""

import os
import sys
import queue
import shutil
import tempfile

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(THIS_DIR)
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

import tkinter as tk
from openpyxl import Workbook

from e_help import utils, history
import e_help.gui as gui
from e_help.excel_reader import read_and_aggregate, _get_excel_files
from e_help.gui import SummaryApp, ExcelFileSelectDialog
from e_help.config import APP_VERSION

n_passed = 0
n_total = 0


def check(name, condition, detail=""):
    global n_passed, n_total
    n_total += 1
    if condition:
        n_passed += 1
        print(f"  PASS  {name}")
    else:
        print(f"  FAIL  {name}  {detail}")


print("=" * 55)
print("  v1.5.2 Hotfix3 — 文件选择机制重构 验证")
print("=" * 55)

# 隔离真实数据目录
TMP = tempfile.mkdtemp(prefix="hotfix3_")
_orig_base_dir = utils.get_base_dir
utils.get_base_dir = lambda: TMP

TEST_DIR = os.path.join(PROJECT_ROOT, "tests", "_hotfix3_test")
if os.path.isdir(TEST_DIR):
    shutil.rmtree(TEST_DIR)
os.makedirs(TEST_DIR)

SCAN_DIR = os.path.join(TEST_DIR, "scan")
os.makedirs(SCAN_DIR)


def make_xlsx(path, headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)
    wb.close()


# 准备扫描目录：3 个有效 Excel + 干扰文件
make_xlsx(os.path.join(SCAN_DIR, "a.xlsx"), ["客户", "金额"], [["公司A", 100.0]])
make_xlsx(os.path.join(SCAN_DIR, "b.xlsx"), ["客户", "金额"], [["公司B", 200.0]])
make_xlsx(os.path.join(SCAN_DIR, "c.xlsm"), ["客户", "金额"], [["公司C", 300.0]])
open(os.path.join(SCAN_DIR, "skip.txt"), "w").close()
open(os.path.join(SCAN_DIR, "~$tmp.xlsx"), "w").close()
open(os.path.join(SCAN_DIR, ".hidden.xlsx"), "w").close()
open(os.path.join(SCAN_DIR, "empty.xlsx"), "w").close()

# 第二个文件夹（用于「重新选择替换」测试）
SCAN2 = os.path.join(TEST_DIR, "scan2")
os.makedirs(SCAN2)
make_xlsx(os.path.join(SCAN2, "d.xlsx"), ["客户", "金额"], [["公司D", 400.0]])


def fresh_app():
    root = tk.Tk()
    root.withdraw()
    try:
        return root, SummaryApp(root)
    except Exception:
        root.destroy()
        raise


# ══════════════════════════════════════════════════════════════
#  T1-T4: read_and_aggregate 的 file_names 参数
# ══════════════════════════════════════════════════════════════
print("\n[read_and_aggregate] file_names 过滤")

# T1: 只汇总指定文件
total, count, fields = read_and_aggregate(SCAN_DIR, SCAN_DIR, file_names=["a.xlsx"])
check("T1a: 只读取 1 个文件", count == 1, f"实际 {count}")
check("T1b: 只含公司A", "公司A" in total and "公司B" not in total and "公司C" not in total,
      f"实际 {list(total.keys())}")

# T2: file_names=None 汇总全部
total2, count2, _ = read_and_aggregate(SCAN_DIR, SCAN_DIR)
check("T2: 汇总全部 3 个文件", count2 == 3 and len(total2) == 3, f"实际 count={count2} keys={list(total2.keys())}")

# T3: file_names=[] 空列表 → 0 文件
total3, count3, _ = read_and_aggregate(SCAN_DIR, SCAN_DIR, file_names=[])
check("T3: 空列表 → 0 文件且 total 为空", count3 == 0 and total3 == {}, f"实际 count={count3} total={total3}")

# T4: file_names 含不存在文件 → 被过滤
total4, count4, _ = read_and_aggregate(SCAN_DIR, SCAN_DIR,
                                       file_names=["a.xlsx", "not_exist.xlsx"])
check("T4: 不存在文件被过滤，只汇总真实文件", count4 == 1 and "公司A" in total4 and len(total4) == 1,
      f"实际 count={count4} keys={list(total4.keys())}")

# ══════════════════════════════════════════════════════════════
#  T5-T11: ExcelFileSelectDialog
# ══════════════════════════════════════════════════════════════
print("\n[ExcelFileSelectDialog] 复选框 / 全选 / 反选 / 计数器")

FILES = ["a.xlsx", "b.xlsx", "c.xlsm"]

root = tk.Tk()
root.withdraw()

# T5: 默认全部勾选
d = ExcelFileSelectDialog(root, SCAN_DIR, FILES)
check("T5: 默认全部勾选", d._selected_names() == FILES, f"实际 {d._selected_names()}")

# T6: 全不选
d._select_none()
check("T6: 全不选后为空", d._selected_names() == [], f"实际 {d._selected_names()}")

# T7: 反选（空 → 全选）
d._invert()
check("T7: 反选后全部勾选", set(d._selected_names()) == set(FILES), f"实际 {d._selected_names()}")

# T8: 全选按钮
d._select_none()
d._select_all()
check("T8: 全选按钮生效", d._selected_names() == FILES, f"实际 {d._selected_names()}")

# T9: 计数器文本
check("T9a: 计数器显示 3/3", d.counter_label.cget("text") == "已选择：3 / 3",
      f"实际 '{d.counter_label.cget('text')}'")
d._select_none()
check("T9b: 计数器更新为 0/3", d.counter_label.cget("text") == "已选择：0 / 3",
      f"实际 '{d.counter_label.cget('text')}'")
d.destroy()

# T10: 空选择确定 → 警告，不关闭，result 仍 None
d2 = ExcelFileSelectDialog(root, SCAN_DIR, ["a.xlsx"])
d2._select_none()
warns = []
_orig_warn = tk.messagebox.showwarning
tk.messagebox.showwarning = lambda title, msg, **kw: warns.append(msg) or "ok"
d2._on_confirm()
tk.messagebox.showwarning = _orig_warn
check("T10a: 空选择确定弹警告", len(warns) > 0 and "请至少选择一个 Excel 文件" in warns[0],
      f"实际 {warns}")
check("T10b: result 仍为 None", d2.result is None, f"实际 {d2.result}")
d2.destroy()

# T11: 有选择确定 → result 为选中列表
d3 = ExcelFileSelectDialog(root, SCAN_DIR, ["a.xlsx", "b.xlsx"])
d3._select_none()
d3._vars["b.xlsx"].set(True)
d3._on_confirm()
check("T11: 确定后返回选中列表", d3.result == ["b.xlsx"], f"实际 {d3.result}")

root.destroy()

# ══════════════════════════════════════════════════════════════
#  T12-T16: _on_folder_selected 选择流程
# ══════════════════════════════════════════════════════════════
print("\n[_on_folder_selected] 扫描 → 弹窗勾选 → 仅勾选文件")

# T12: 空文件夹 → 提示，不弹窗，禁用
empty_dir = os.path.join(TEST_DIR, "empty")
os.makedirs(empty_dir)
root, app = fresh_app()
called = []
app._prompt_select_files = lambda folder, files: called.append((folder, files)) or None
app.selected_folder.set(empty_dir)
app._on_folder_selected()
check("T12a: 提示无 Excel 文件", "当前文件夹中没有可汇总的 Excel 文件" in app.file_status_label.cget("text"),
      f"实际 '{app.file_status_label.cget('text')}'")
check("T12b: 未弹出选择对话框", len(called) == 0, f"实际调用 {len(called)} 次")
check("T12c: selected_excel_files 为空", app.selected_excel_files == [], f"实际 {app.selected_excel_files}")
check("T12d: 开始按钮禁用", str(app.start_btn.cget("state")) == "disabled",
      f"实际 {app.start_btn.cget('state')}")
root.destroy()

# T13: 正常文件夹 → 弹窗传全部文件，selected_excel_files 只含勾选
root, app = fresh_app()
captured = {}
def fake_prompt(folder, files):
    captured["folder"] = folder
    captured["files"] = list(files)
    return ["a.xlsx", "b.xlsx"]
app._prompt_select_files = fake_prompt
app.selected_folder.set(SCAN_DIR)
app._on_folder_selected()
check("T13a: 传入正确文件夹", captured.get("folder") == SCAN_DIR, f"实际 {captured.get('folder')}")
check("T13b: 传入全部有效 Excel（已排序）",
      captured.get("files") == sorted(["a.xlsx", "b.xlsx", "c.xlsm"]), f"实际 {captured.get('files')}")
check("T13c: selected_excel_files 只含勾选", app.selected_excel_files == ["a.xlsx", "b.xlsx"],
      f"实际 {app.selected_excel_files}")
check("T13d: 只扫描勾选文件", set(app._preview_ok) == {"a.xlsx", "b.xlsx"},
      f"实际 ok={app._preview_ok} warn={app._preview_warn}")
root.destroy()

# T14: 取消 → 空选择，禁用
root, app = fresh_app()
app._prompt_select_files = lambda folder, files: None
app.selected_folder.set(SCAN_DIR)
app._on_folder_selected()
check("T14a: 取消后 selected_excel_files 为空", app.selected_excel_files == [],
      f"实际 {app.selected_excel_files}")
check("T14b: 开始按钮禁用", str(app.start_btn.cget("state")) == "disabled",
      f"实际 {app.start_btn.cget('state')}")
root.destroy()

# T15: 重新选择文件夹 → 替换旧列表，不混合
root, app = fresh_app()
app._prompt_select_files = lambda folder, files: ["a.xlsx"]
app.selected_folder.set(SCAN_DIR)
app._on_folder_selected()
check("T15a: 第一次勾选 a.xlsx", app.selected_excel_files == ["a.xlsx"],
      f"实际 {app.selected_excel_files}")
app._prompt_select_files = lambda folder, files: ["d.xlsx"]
app.selected_folder.set(SCAN2)
app._on_folder_selected()
check("T15b: 重新选择后替换，不含旧文件", app.selected_excel_files == ["d.xlsx"],
      f"实际 {app.selected_excel_files}")
check("T15c: _preview_ok 不含 a.xlsx", "a.xlsx" not in app._preview_ok,
      f"实际 {app._preview_ok}")
root.destroy()

# T16: _get_excel_files 过滤干扰文件
files = _get_excel_files(SCAN_DIR)
check("T16a: 只含 3 个有效 Excel", sorted(files) == sorted(["a.xlsx", "b.xlsx", "c.xlsm"]),
      f"实际 {files}")
check("T16b: 干扰文件被过滤",
      "skip.txt" not in files and "~$tmp.xlsx" not in files
      and ".hidden.xlsx" not in files and "empty.xlsx" not in files, f"实际 {files}")

# ══════════════════════════════════════════════════════════════
#  T17-T19: _on_start_summary 只使用 selected_excel_files
# ══════════════════════════════════════════════════════════════
print("\n[_on_start_summary] 只使用勾选文件")


class FakeWorker:
    """替身工作线程：记录构造参数，不真正执行。"""
    instances = []

    def __init__(self, folder, output_path, base_dir, file_names=None, field_mappings=None):
        self.folder = folder
        self.output_path = output_path
        self.base_dir = base_dir
        self.file_names = file_names
        self.field_mappings = field_mappings
        self.queue = queue.Queue()
        FakeWorker.instances.append(self)

    def start(self):
        pass

    def is_alive(self):
        return False


_orig_worker = gui._WorkerThread
gui._WorkerThread = FakeWorker

# T17: 空选择 → 警告，worker 未启动
root, app = fresh_app()
app.selected_folder.set(SCAN_DIR)
app.selected_excel_files = []
warns = []
_orig_warn2 = gui.messagebox.showwarning
gui.messagebox.showwarning = lambda title, msg, **kw: warns.append(msg) or "ok"
app._on_start_summary()
gui.messagebox.showwarning = _orig_warn2
check("T17a: 空选择弹警告", len(warns) > 0 and "请先选择要汇总的 Excel 文件" in warns[0],
      f"实际 {warns}")
check("T17b: worker 未启动", app.worker is None, f"实际 {app.worker}")
root.destroy()

# T18: 正常启动 → 只汇总勾选文件
root, app = fresh_app()
app.selected_folder.set(SCAN_DIR)
app.selected_excel_files = ["a.xlsx", "c.xlsm"]
app._on_start_summary()
check("T18a: _task_input_files 只含勾选文件", app._task_input_files == ["a.xlsx", "c.xlsm"],
      f"实际 {app._task_input_files}")
check("T18b: worker.file_names 只含勾选文件",
      isinstance(app.worker, FakeWorker) and app.worker.file_names == ["a.xlsx", "c.xlsm"],
      f"实际 {app.worker.file_names if app.worker else None}")
check("T18c: 未扫描整个文件夹（b.xlsx 不在列表）", "b.xlsx" not in app._task_input_files,
      f"实际 {app._task_input_files}")
root.destroy()

# T19: _on_start_new_task 清空选择
root, app = fresh_app()
app.selected_excel_files = ["a.xlsx"]
app._preview_ok.append("a.xlsx")
app._on_start_new_task()
check("T19: 新建任务清空选择与预览", app.selected_excel_files == [] and app._preview_ok == [],
      f"实际 sel={app.selected_excel_files} ok={app._preview_ok}")
root.destroy()

gui._WorkerThread = _orig_worker

# ══════════════════════════════════════════════════════════════
#  T20-T21: 字段映射 / 字段规则 只对已勾选文件重新识别
# ══════════════════════════════════════════════════════════════
print("\n[字段映射 / 字段规则] 只重新识别已勾选文件")


def make_auto_close_dialog(result):
    class _Fake(tk.Toplevel):
        def __init__(self, parent, *args, **kwargs):
            super().__init__(parent)
            self.withdraw()
            self.result = result
            self.mappings = {}
            self.after(0, self.destroy)
    return _Fake


# T20: 字段映射后调用 _rescan_selected_files（有选择时）
root, app = fresh_app()
app.selected_folder.set(SCAN_DIR)
app.selected_excel_files = ["a.xlsx"]
calls = {"rescan": 0, "reselect": 0}
app._rescan_selected_files = lambda: calls.__setitem__("rescan", calls["rescan"] + 1)
app._on_folder_selected = lambda: calls.__setitem__("reselect", calls["reselect"] + 1)
_orig_fmd = gui.FieldMappingDialog
gui.FieldMappingDialog = make_auto_close_dialog(True)
app._on_open_mapping_dialog("a.xlsx")
gui.FieldMappingDialog = _orig_fmd
check("T20a: 字段映射后调用 _rescan_selected_files", calls["rescan"] == 1, f"实际 {calls}")
check("T20b: 未重新弹出文件夹选择", calls["reselect"] == 0, f"实际 {calls}")
root.destroy()

# T21: 字段规则后调用 _rescan_selected_files（有选择时）
root, app = fresh_app()
app.selected_folder.set(SCAN_DIR)
app.selected_excel_files = ["a.xlsx"]
calls2 = {"rescan": 0, "reselect": 0}
app._rescan_selected_files = lambda: calls2.__setitem__("rescan", calls2["rescan"] + 1)
app._on_folder_selected = lambda: calls2.__setitem__("reselect", calls2["reselect"] + 1)
_orig_fcd = gui.FieldConfigDialog
gui.FieldConfigDialog = make_auto_close_dialog(True)
app._on_open_field_config()
gui.FieldConfigDialog = _orig_fcd
check("T21a: 字段规则后调用 _rescan_selected_files", calls2["rescan"] == 1, f"实际 {calls2}")
check("T21b: 未重新弹出文件夹选择", calls2["reselect"] == 0, f"实际 {calls2}")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T22-T24: 版本 / 端到端 / 全选保持旧行为
# ══════════════════════════════════════════════════════════════
print("\n[版本 / 端到端]")

# T22: 版本号
check("T22: 版本号 = v1.5.4", APP_VERSION == "v1.5.4", f"实际 {APP_VERSION}")

# T23: 端到端 — 只汇总勾选文件（跨文件求和）
make_xlsx(os.path.join(SCAN_DIR, "e.xlsx"), ["客户", "金额"], [["公司A", 50.0]])
total_e, count_e, _ = read_and_aggregate(SCAN_DIR, SCAN_DIR, file_names=["a.xlsx", "e.xlsx"])
check("T23a: 只汇总勾选的 2 个文件", count_e == 2, f"实际 {count_e}")
check("T23b: 公司A 跨文件求和 = 150", total_e.get("公司A", {}).get("金额") == 150.0,
      f"实际 {total_e.get('公司A')}")
check("T23c: 公司B/公司C 未汇总", "公司B" not in total_e and "公司C" not in total_e,
      f"实际 {list(total_e.keys())}")

# T24: 全选（默认勾选全部）结果与全文件夹一致，保持旧行为
all_files = sorted(_get_excel_files(SCAN_DIR))
total_sel, count_sel, _ = read_and_aggregate(SCAN_DIR, SCAN_DIR, file_names=all_files)
total_all, count_all, _ = read_and_aggregate(SCAN_DIR, SCAN_DIR)
check("T24: 全选结果与全文件夹一致", count_sel == count_all and total_sel == total_all,
      f"全选 count={count_sel} vs 全部 count={count_all}")

# 清理
history.clear_records()
utils.get_base_dir = _orig_base_dir
shutil.rmtree(TEST_DIR, ignore_errors=True)
shutil.rmtree(TMP, ignore_errors=True)

print("\n" + "=" * 55)
print(f"  {n_passed}/{n_total} 通过")
print("=" * 55)
if n_passed != n_total:
    sys.exit(1)
