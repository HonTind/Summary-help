"""
v1.5.2 Hotfix2 — 恢复 Windows 原生文件夹选择器 验证测试

  T1-T5  : 原生选择器（filedialog.askdirectory）行为
  T6-T7  : 扫描 / 字段识别 / 动态字段
  T8     : 任务记录
  T9-T10 : 自定义输出目录 / 重启
  T11    : 版本统一

运行:
  cd "Day07(小试牛刀)源文件"
  python -m Excel自动汇总助手.tests.verify_hotfix2
"""

import os
import sys
import json
import types
import shutil
import tempfile

# TEST-PHONE-* 是字段保留测试的合成占位内容，不是联系电话。
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(THIS_DIR)
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

import tkinter as tk
from openpyxl import Workbook, load_workbook
import xlwt

from Excel自动汇总助手 import utils, history
import Excel自动汇总助手.gui as gui
from Excel自动汇总助手.excel_reader import read_and_aggregate
from Excel自动汇总助手.field_mapper import (
    detect_columns, save_field_config, reset_field_config,
)
from Excel自动汇总助手.gui import SummaryApp
from Excel自动汇总助手.config import APP_VERSION

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
print("  v1.5.2 Hotfix2 — 恢复 Windows 原生选择器 验证")
print("=" * 55)

# 隔离真实数据目录
TMP = tempfile.mkdtemp(prefix="hotfix2_")
_orig_base_dir = utils.get_base_dir
utils.get_base_dir = lambda: TMP

TEST_DIR = os.path.join(PROJECT_ROOT, "tests", "_hotfix2_test")
if os.path.isdir(TEST_DIR):
    shutil.rmtree(TEST_DIR)
os.makedirs(TEST_DIR)

# 准备含 Excel 文件的扫描目录
SCAN_DIR = os.path.join(TEST_DIR, "scan")
os.makedirs(SCAN_DIR)

wb = Workbook()
ws = wb.active
ws.append(["客户", "金额"])
ws.append(["公司A", 100.0])
wb.save(os.path.join(SCAN_DIR, "a.xlsx"))
wb.close()

wb2 = xlwt.Workbook()
sh2 = wb2.add_sheet("S")
sh2.write(0, 0, "客户")
sh2.write(0, 1, "金额")
sh2.write(1, 0, "公司B")
sh2.write(1, 1, 200.0)
wb2.save(os.path.join(SCAN_DIR, "b.xls"))

wb3 = Workbook()
ws3 = wb3.active
ws3.append(["客户", "金额"])
ws3.append(["公司C", 300.0])
wb3.save(os.path.join(SCAN_DIR, "c.xlsm"))
wb3.close()

# 干扰文件（应被过滤）
open(os.path.join(SCAN_DIR, "skip.txt"), "w").close()
open(os.path.join(SCAN_DIR, "~$tmp.xlsx"), "w").close()
open(os.path.join(SCAN_DIR, ".hidden.xlsx"), "w").close()


def fresh_app():
    root = tk.Tk()
    root.withdraw()
    try:
        return root, SummaryApp(root)
    except Exception:
        root.destroy()
        raise


# ══════════════════════════════════════════════════════════════
#  T1-T5: 原生选择器行为
# ══════════════════════════════════════════════════════════════
print("\n[原生选择器] filedialog.askdirectory")

# T1: _on_browse_folder 调用原生 askdirectory（非自制对话框）
print("\n[T1] _on_browse_folder 使用原生 askdirectory")
calls = {"n": 0, "kwargs": None}
orig_filedialog = gui.filedialog


def fake_askdirectory(**kwargs):
    calls["n"] += 1
    calls["kwargs"] = kwargs
    return ""


gui.filedialog = types.SimpleNamespace(askdirectory=fake_askdirectory)
root, app = fresh_app()
app._on_browse_folder()
check("T1a: 调用了 askdirectory", calls["n"] == 1, f"实际调用次数 {calls['n']}")
check("T1b: 传入标题参数", "title" in (calls["kwargs"] or {}), f"实际 {calls['kwargs']}")
root.destroy()

# T2: 用户取消（返回空字符串）→ 不崩溃、路径不变
print("\n[T2] 取消选择不崩溃")
gui.filedialog = types.SimpleNamespace(askdirectory=lambda **kw: "")
root, app = fresh_app()
app.selected_folder.set("之前的值")
app._on_browse_folder()
check("T2: 取消后路径不变", app.selected_folder.get() == "之前的值",
      f"实际 '{app.selected_folder.get()}'")
root.destroy()

# T3: 选择有效文件夹 → 路径正确返回并触发扫描
print("\n[T3] 选择有效文件夹")
gui.filedialog = types.SimpleNamespace(askdirectory=lambda **kw: SCAN_DIR)
root, app = fresh_app()
app._prompt_select_files = lambda folder, files: list(files)
app._on_browse_folder()
check("T3: 路径已设置", app.selected_folder.get() == SCAN_DIR,
      f"实际 '{app.selected_folder.get()}'")
root.destroy()

# T4: .xlsx / .xls / .xlsm 均被扫描，干扰文件被过滤
print("\n[T4] 扫描 .xlsx/.xls/.xlsm")
gui.filedialog = types.SimpleNamespace(askdirectory=lambda **kw: SCAN_DIR)
root, app = fresh_app()
app._prompt_select_files = lambda folder, files: list(files)
app._on_browse_folder()
ok = set(app._preview_ok)
check("T4a: a.xlsx 被识别", "a.xlsx" in ok, f"实际 {ok}")
check("T4b: b.xls 被识别", "b.xls" in ok, f"实际 {ok}")
check("T4c: c.xlsm 被识别", "c.xlsm" in ok, f"实际 {ok}")
check("T4d: 干扰文件被过滤", "skip.txt" not in ok and not any(f.startswith("~$") for f in ok),
      f"实际 {ok}")
root.destroy()

# T5: 文件列表正常显示（preview_ok 已填充且无 warning）
print("\n[T5] 文件列表正常显示")
gui.filedialog = types.SimpleNamespace(askdirectory=lambda **kw: SCAN_DIR)
root, app = fresh_app()
app._prompt_select_files = lambda folder, files: list(files)
app._on_browse_folder()
check("T5a: 3 个文件进入预览", len(app._preview_ok) == 3, f"实际 {app._preview_ok}")
check("T5b: 无 warning 文件", len(app._preview_warn) == 0, f"实际 {app._preview_warn}")
root.destroy()

gui.filedialog = orig_filedialog

# ══════════════════════════════════════════════════════════════
#  T6-T7: 字段识别 / 动态字段
# ══════════════════════════════════════════════════════════════
print("\n[字段识别 / 动态字段]")

# T6: 默认字段识别（客户/金额）
print("\n[T6] 默认字段识别")
wb = load_workbook(os.path.join(SCAN_DIR, "a.xlsx"), data_only=True)
det = detect_columns(wb.active)
wb.close()
names = [d.field_name for d in (det or [])]
check("T6: 客户+金额 识别成功", "客户" in names and "金额" in names, f"实际 {names}")

# T7: 动态字段（3 字段）聚合正常
print("\n[T7] 动态字段（3 字段）")
reset_field_config(TEST_DIR)
save_field_config(TEST_DIR, [
    {"name": "客户", "keywords": ["客户", "公司"]},
    {"name": "金额", "keywords": ["金额", "总金额"]},
    {"name": "电话", "keywords": ["电话", "联系电话"]},
])
wb = Workbook()
ws = wb.active
ws.append(["客户", "金额", "电话"])
ws.append(["公司A", 100.0, "TEST-PHONE-02"])
ws.append(["公司B", 200.0, "TEST-PHONE-03"])
wb.save(os.path.join(TEST_DIR, "dyn.xlsx"))
wb.close()
total, count, fields = read_and_aggregate(TEST_DIR, TEST_DIR)
check("T7a: 3 字段顺序正确", fields == ["客户", "金额", "电话"], f"实际 {fields}")
check("T7b: 客户全部保留", len(total) == 2, f"实际 {list(total.keys())}")
check("T7c: 电话字段保留", total.get("公司A", {}).get("电话") == "TEST-PHONE-02",
      f"实际 {total.get('公司A')}")
os.remove(os.path.join(TEST_DIR, "dyn.xlsx"))
reset_field_config(TEST_DIR)

# ══════════════════════════════════════════════════════════════
#  T8: 任务记录
# ══════════════════════════════════════════════════════════════
print("\n[任务记录]")
history.clear_records()
history.add_record({"status": "success", "file_count": 3, "fields": ["客户", "金额"]})
recs = history.load_history()
check("T8a: 记录已写入", len(recs) == 1, f"实际 {len(recs)}")
check("T8b: 状态正确", recs[0].get("status") == "success", f"实际 {recs[0].get('status')}")
history.clear_records()

# ══════════════════════════════════════════════════════════════
#  T9-T10: 自定义输出目录 / 重启
# ══════════════════════════════════════════════════════════════
print("\n[自定义输出目录 / 重启]")
config_dir = utils.get_config_dir()
out_dir = os.path.join(TMP, "输出目录")
os.makedirs(out_dir, exist_ok=True)
with open(os.path.join(config_dir, "settings.json"), "w", encoding="utf-8") as f:
    json.dump({"output_dir": out_dir}, f, ensure_ascii=False)

# T9: 首次实例化加载自定义输出目录
print("\n[T9] 加载自定义输出目录")
root, app = fresh_app()
check("T9: 自定义目录已加载", app.custom_output_dir.get() == out_dir,
      f"实际 '{app.custom_output_dir.get()}'")
root.destroy()

# T10: 重启（再次实例化）仍加载
print("\n[T10] 重启后仍加载")
root, app = fresh_app()
check("T10: 重启后仍加载同一目录", app.custom_output_dir.get() == out_dir,
      f"实际 '{app.custom_output_dir.get()}'")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T11: 版本统一
# ══════════════════════════════════════════════════════════════
print("\n[版本]")

check("T11: 版本号 = v1.5.4", APP_VERSION == "v1.5.4",
      f"实际 {APP_VERSION}")

# 清理
utils.get_base_dir = _orig_base_dir
reset_field_config(TEST_DIR)
shutil.rmtree(TEST_DIR, ignore_errors=True)
shutil.rmtree(TMP, ignore_errors=True)

print("\n" + "=" * 55)
print(f"  {n_passed}/{n_total} 通过")
print("=" * 55)
if n_passed != n_total:
    sys.exit(1)
