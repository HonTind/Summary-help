"""
v1.5.3 — 输出目录持久化 + 内存压力测试 验证测试

  T1-T10  : 输出目录持久化
             T1  默认输出模式正常
             T2  输入文件夹输出模式正常
             T3  自定义输出目录正常
             T4  选择自定义目录后重新选择输入文件夹，输出目录不变
             T5  自定义目录保存到 settings.json
             T6  重启后自定义目录恢复
             T7  重新选择多个输入文件夹，输出目录始终保持
             T8  用户主动修改输出目录后使用新目录
             T9  切换到默认模式后正确恢复默认输出
             T10 切换模式不会产生错误残留
  T11-T14 : 内存压力测试自动化入口（100/300/500/1000 文件）
  T15     : stress_test_report.txt 生成
  T16     : 版本 = v1.5.3

运行:
  cd "Day07(小试牛刀)源文件"
  python -m e_help.tests.verify_v153
"""

import os
import sys
import json
import types
import queue as _queue_module
import shutil
import tempfile

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(THIS_DIR)
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

import tkinter as tk
from openpyxl import Workbook

from e_help import utils
import e_help.gui as gui
from e_help.gui import SummaryApp
from e_help.config import (
    APP_NAME, APP_VERSION, APP_FULL_NAME, OUTPUT_FOLDER_NAME,
)
from e_help.tests import stress_test_v153 as stress

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
print("  v1.5.3 输出目录持久化 + 内存压力测试 验证")
print("=" * 55)

# 独立测试环境：把 get_base_dir 指到临时目录，避免污染真实数据目录。
# 注意：gui.py 通过 `from .utils import get_base_dir` 引用了原函数，
#       因此需要同时打补丁 gui.get_base_dir，否则默认输出目录会落到真实项目目录。
TMP = tempfile.mkdtemp(prefix="v153_")
_orig_utils_base_dir = utils.get_base_dir
_orig_gui_base_dir = gui.get_base_dir
utils.get_base_dir = lambda: TMP
gui.get_base_dir = lambda: TMP

_orig_filedialog = gui.filedialog
_orig_worker = gui._WorkerThread

# 测试目录
SCAN1 = os.path.join(TMP, "输入A")
SCAN2 = os.path.join(TMP, "输入B")
CUSTOM1 = os.path.join(TMP, "我的输出1")
CUSTOM2 = os.path.join(TMP, "我的输出2")
for d in (SCAN1, SCAN2, CUSTOM1, CUSTOM2):
    os.makedirs(d, exist_ok=True)


def make_xlsx(path):
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["公司A", 100.0])
    wb.save(path)
    wb.close()


make_xlsx(os.path.join(SCAN1, "a.xlsx"))
make_xlsx(os.path.join(SCAN2, "b.xlsx"))


def _settings_path():
    return os.path.join(utils.get_config_dir(), "settings.json")


def _clear_settings():
    p = _settings_path()
    if os.path.isfile(p):
        os.remove(p)


def fresh_app():
    root = tk.Tk()
    root.withdraw()
    try:
        return root, SummaryApp(root)
    except Exception:
        root.destroy()
        raise


class FakeWorker:
    """替换 _WorkerThread，让 _on_start_summary 同步完成、不读真实文件。"""

    def __init__(self, folder, output_path, base_dir, file_names=None, field_mappings=None):
        self.folder = folder
        self.output_path = output_path
        self.base_dir = base_dir
        self.file_names = file_names
        self.field_mappings = field_mappings
        self.queue = _queue_module.Queue()

    def start(self):
        pass

    def is_alive(self):
        return False


gui._WorkerThread = FakeWorker


def set_custom(app, directory):
    """将输出模式设为 custom 并指定目录（模拟用户勾选自定义目录）。"""
    app.output_mode.set("custom")
    app.custom_output_dir.set(directory)
    app._save_output_setting()


# ══════════════════════════════════════════════════════════════
#  T1: 默认输出模式正常
# ══════════════════════════════════════════════════════════════
print("\n[T1] 默认输出模式")
_clear_settings()
root, app = fresh_app()
check("T1a: 初始 output_mode = default", app.output_mode.get() == "default",
      f"实际 {app.output_mode.get()!r}")
check("T1b: 初始自定义目录为空", app.custom_output_dir.get().strip() == "",
      f"实际 {app.custom_output_dir.get()!r}")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T2: 输入文件夹输出模式正常
# ══════════════════════════════════════════════════════════════
print("\n[T2] 输入文件夹输出模式")
_clear_settings()
root, app = fresh_app()
app.output_mode.set("input")
app.selected_folder.set(SCAN1)
app.selected_excel_files = ["a.xlsx"]
app._on_start_summary()
check("T2: input 模式输出到输入文件夹", os.path.dirname(app.output_path) == SCAN1,
      f"实际 {app.output_path!r}")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T3: 自定义输出目录正常
# ══════════════════════════════════════════════════════════════
print("\n[T3] 自定义输出目录")
_clear_settings()
root, app = fresh_app()
set_custom(app, CUSTOM1)
app.selected_folder.set(SCAN1)
app.selected_excel_files = ["a.xlsx"]
app._on_start_summary()
check("T3: custom 模式输出到自定义目录", os.path.dirname(app.output_path) == CUSTOM1,
      f"实际 {app.output_path!r}")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T4: 选择自定义目录后重新选择输入文件夹，输出目录不变
# ══════════════════════════════════════════════════════════════
print("\n[T4] 重新选择输入文件夹不重置输出目录")
_clear_settings()
gui.filedialog = types.SimpleNamespace(askdirectory=lambda **kw: SCAN1)
root, app = fresh_app()
set_custom(app, CUSTOM1)
app._prompt_select_files = lambda folder, files: list(files)
app._on_browse_folder()          # 完整走「浏览 → 选择文件夹」路径
check("T4a: 重选后 output_mode 仍为 custom", app.output_mode.get() == "custom",
      f"实际 {app.output_mode.get()!r}")
check("T4b: 重选后自定义目录不变", app.custom_output_dir.get() == CUSTOM1,
      f"实际 {app.custom_output_dir.get()!r}")
check("T4c: 输入目录已更新", app.selected_folder.get() == SCAN1,
      f"实际 {app.selected_folder.get()!r}")
root.destroy()

# 新建任务（_on_start_new_task）也不重置输出
_clear_settings()
root, app = fresh_app()
set_custom(app, CUSTOM1)
app._on_start_new_task()
check("T4d: 新建任务后 output_mode 仍为 custom", app.output_mode.get() == "custom",
      f"实际 {app.output_mode.get()!r}")
check("T4e: 新建任务后自定义目录不变", app.custom_output_dir.get() == CUSTOM1,
      f"实际 {app.custom_output_dir.get()!r}")
root.destroy()
gui.filedialog = _orig_filedialog

# ══════════════════════════════════════════════════════════════
#  T5: 自定义目录保存到 settings.json
# ══════════════════════════════════════════════════════════════
print("\n[T5] 自定义目录保存到 settings.json")
_clear_settings()
root, app = fresh_app()
set_custom(app, CUSTOM1)
with open(_settings_path(), "r", encoding="utf-8") as f:
    s = json.load(f)
check("T5a: settings.json 含 output_mode=custom", s.get("output_mode") == "custom",
      f"实际 {s.get('output_mode')!r}")
check("T5b: settings.json 含 output_dir", s.get("output_dir") == CUSTOM1,
      f"实际 {s.get('output_dir')!r}")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T6: 重启后自定义目录恢复
# ══════════════════════════════════════════════════════════════
print("\n[T6] 重启后自定义目录恢复")
_clear_settings()
root, app = fresh_app()
set_custom(app, CUSTOM1)
root.destroy()

root, app = fresh_app()          # 模拟重启：重新实例化
check("T6a: 重启后 output_mode 恢复 custom", app.output_mode.get() == "custom",
      f"实际 {app.output_mode.get()!r}")
check("T6b: 重启后自定义目录恢复", app.custom_output_dir.get() == CUSTOM1,
      f"实际 {app.custom_output_dir.get()!r}")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T7: 重新选择多个输入文件夹，输出目录始终保持
# ══════════════════════════════════════════════════════════════
print("\n[T7] 重新选择多个输入文件夹")
_clear_settings()
gui.filedialog = types.SimpleNamespace(askdirectory=lambda **kw: SCAN1)
root, app = fresh_app()
set_custom(app, CUSTOM1)
app._prompt_select_files = lambda folder, files: list(files)
app._on_browse_folder()
check("T7a: 第 1 次重选后目录不变", app.custom_output_dir.get() == CUSTOM1,
      f"实际 {app.custom_output_dir.get()!r}")
gui.filedialog = types.SimpleNamespace(askdirectory=lambda **kw: SCAN2)
app._on_browse_folder()
check("T7b: 第 2 次重选后目录不变", app.custom_output_dir.get() == CUSTOM1,
      f"实际 {app.custom_output_dir.get()!r}")
check("T7c: 输入文件夹已切换到 SCAN2", app.selected_folder.get() == SCAN2,
      f"实际 {app.selected_folder.get()!r}")
root.destroy()
gui.filedialog = _orig_filedialog

# ══════════════════════════════════════════════════════════════
#  T8: 用户主动修改输出目录后使用新目录
# ══════════════════════════════════════════════════════════════
print("\n[T8] 用户主动修改输出目录后使用新目录")
_clear_settings()
root, app = fresh_app()
set_custom(app, CUSTOM1)
app.custom_output_dir.set(CUSTOM2)   # 用户主动改到 CUSTOM2
app._save_output_setting()
app.selected_folder.set(SCAN1)
app.selected_excel_files = ["a.xlsx"]
app._on_start_summary()
check("T8: 使用新目录 CUSTOM2", os.path.dirname(app.output_path) == CUSTOM2,
      f"实际 {app.output_path!r}")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T9: 切换到默认模式后正确恢复默认输出
# ══════════════════════════════════════════════════════════════
print("\n[T9] 切换到默认模式后恢复默认输出")
_clear_settings()
root, app = fresh_app()
set_custom(app, CUSTOM1)
app.output_mode.set("default")       # 切回默认模式
app.selected_folder.set(SCAN1)
app.selected_excel_files = ["a.xlsx"]
app._on_start_summary()
default_dir = os.path.join(TMP, OUTPUT_FOLDER_NAME)
check("T9: 默认模式输出到默认目录", os.path.dirname(app.output_path) == default_dir,
      f"实际 {app.output_path!r} 期望 {default_dir!r}")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T10: 切换模式不会产生错误残留
# ══════════════════════════════════════════════════════════════
print("\n[T10] 切换模式不产生错误残留")
_clear_settings()
root, app = fresh_app()
set_custom(app, CUSTOM1)
app.output_mode.set("default")
app.output_mode.set("custom")
check("T10a: 切回 custom 后目录仍保留", app.custom_output_dir.get() == CUSTOM1,
      f"实际 {app.custom_output_dir.get()!r}")
check("T10b: 模式为 custom", app.output_mode.get() == "custom",
      f"实际 {app.output_mode.get()!r}")
with open(_settings_path(), "r", encoding="utf-8") as f:
    s = json.load(f)
check("T10c: settings 中 output_mode 为 custom", s.get("output_mode") == "custom",
      f"实际 {s.get('output_mode')!r}")
root.destroy()

# 输出目录测试结束，恢复 worker 引用
gui._WorkerThread = _orig_worker

# ══════════════════════════════════════════════════════════════
#  T11-T14: 内存压力测试（100/300/500/1000 文件）
# ══════════════════════════════════════════════════════════════
print("\n[T11-T14] 内存压力测试（100/300/500/1000 文件）")
stress_root = tempfile.mkdtemp(prefix="v153_stress_")
results = stress.run_full_stress(stress_root=stress_root)
shutil.rmtree(stress_root, ignore_errors=True)

for i, n in enumerate([100, 300, 500, 1000]):
    r = results[i]
    ok = (r["n_files"] == n) and r["success"] and r["correct"]
    check(f"T{11 + i}: {n} 文件压力测试通过", ok,
          f"n={r['n_files']} success={r['success']} correct={r['correct']} "
          f"detail={r['detail']} error={r['error']}")

# T15: 生成 stress_test_report.txt
print("\n[T15] stress_test_report.txt 生成")
report = stress.build_report(results)
with open(stress.REPORT_PATH, "w", encoding="utf-8") as f:
    f.write(report + "\n")
check("T15: stress_test_report.txt 已生成", os.path.isfile(stress.REPORT_PATH),
      f"路径 {stress.REPORT_PATH}")

# T16: 版本统一
print("\n[T16] 版本统一")
check("T16a: APP_VERSION == v1.5.4", APP_VERSION == "v1.5.4", f"实际 {APP_VERSION}")
check("T16b: APP_FULL_NAME 含 v1.5.4", "v1.5.4" in APP_FULL_NAME, f"实际 {APP_FULL_NAME}")
check("T16c: APP_NAME == E-help", APP_NAME == "E-help",
      f"实际 {APP_NAME}")

# ══════════════════════════════════════════════════════════════
#  清理
# ══════════════════════════════════════════════════════════════
utils.get_base_dir = _orig_utils_base_dir
gui.get_base_dir = _orig_gui_base_dir
gui.filedialog = _orig_filedialog
gui._WorkerThread = _orig_worker
shutil.rmtree(TMP, ignore_errors=True)

print("\n" + "=" * 55)
print(f"  {n_passed}/{n_total} 通过")
print("=" * 55)
if n_passed != n_total:
    sys.exit(1)
