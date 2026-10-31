"""
v1.5.4 — 字段映射「临时生效」+ 汇总完整性 验证测试

覆盖 T1-T16：

  T1   正常表头自动识别
  T2   「ccc」等异常表头无法自动识别
  T3   手动映射 + 不勾选「记住」，当前任务仍能成功汇总
  T4   汇总输出包含数据
  T5   客户数量正确
  T6   总金额正确
  T7   可选字段（电话）缺失不丢失客户数据
  T8   勾选「记住」写入 settings.json
  T9   不勾选「记住」不写 settings.json
  T10  不记住 + 重启后临时映射消失
  T11  多个文件中一个异常（按文件/表头上下文映射）
  T12  核心字段无法确认 → 阻止生成汇总
  T13  错误信息清晰（含文件名 + 缺少字段 + 引导）
  T14  可选字段缺失允许继续汇总
  T15  取消不改变原映射
  T16  连续两个任务互不污染

运行:
  cd "Day07(小试牛刀)源文件"
  python -m e_help.tests.verify_v154
"""

import os
import sys
import json
import shutil
import tempfile

# TEST-PHONE-* 是字段保留测试的合成占位内容，不是联系电话。
THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(THIS_DIR)
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

import tkinter as tk
from openpyxl import Workbook

from e_help import utils
import e_help.gui as gui
import e_help.field_mapper as fm
from e_help.gui import SummaryApp
from e_help.field_mapper import detect_columns, save_field_config, reset_field_config
from e_help.excel_reader import aggregate_with_report, read_and_aggregate
from e_help.config import APP_NAME, APP_VERSION, APP_FULL_NAME

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


print("=" * 60)
print("  v1.5.4 字段映射临时生效 + 汇总完整性 验证")
print("=" * 60)

# 独立测试环境：把 get_base_dir 指到临时目录，避免污染真实数据目录。
TMP = tempfile.mkdtemp(prefix="v154_")
_orig_utils_base_dir = utils.get_base_dir
_orig_gui_base_dir = gui.get_base_dir
utils.get_base_dir = lambda: TMP
gui.get_base_dir = lambda: TMP

config_dir = utils.get_config_dir()
logs_dir = utils.get_logs_dir()

# 3 字段配置（客户 + 金额 + 电话），电话作为可选字段用于 T7/T14
FIELD_CONFIG_3 = [
    {"name": "客户", "keywords": ["客户", "客户名称", "购货单位", "单位名称", "公司", "顾客", "收货人", "买方"]},
    {"name": "金额", "keywords": ["金额", "总金额", "合计金额", "价税合计", "货款", "费用", "小计", "应收金额"]},
    {"name": "电话", "keywords": ["电话", "联系电话", "手机", "联系方式"]},
]
save_field_config(config_dir, FIELD_CONFIG_3)

# 快照 CUSTOM_HEADER_MAP，测试结束后恢复，防止污染
_orig_custom_map = dict(fm.CUSTOM_HEADER_MAP)

FOLDER = os.path.join(TMP, "输入")
os.makedirs(FOLDER, exist_ok=True)


def make_xlsx(path, headers, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)
    wb.close()


# 测试文件
A_PATH = os.path.join(FOLDER, "a.xlsx")   # 正常：客户|金额|电话
B_PATH = os.path.join(FOLDER, "b.xlsx")   # 异常：ccc|金额|电话（ccc → 客户）
C_PATH = os.path.join(FOLDER, "c.xlsx")   # 异常：对手方|金额（对手方 → 客户，无电话）
D_PATH = os.path.join(FOLDER, "d.xlsx")   # 异常：xxx|yyy（两个核心字段都无法识别）
E_PATH = os.path.join(FOLDER, "e.xlsx")   # 正常但缺可选字段：客户|金额
MAP_PATH = os.path.join(FOLDER, "map.xlsx")  # 用于对话框测试：ccc|金额

make_xlsx(A_PATH, ["客户", "金额", "电话"], [["公司A", 100, "TEST-PHONE-02"]])
make_xlsx(B_PATH, ["ccc", "金额", "电话"], [["公司B", 200, "TEST-PHONE-03"]])
make_xlsx(C_PATH, ["对手方", "金额"], [["公司C", 50]])
make_xlsx(D_PATH, ["xxx", "yyy"], [["公司D", 1]])
make_xlsx(E_PATH, ["客户", "金额"], [["公司E", 30]])
make_xlsx(MAP_PATH, ["ccc", "金额"], [["公司X", 5]])


def _settings_path():
    return os.path.join(config_dir, "settings.json")


def _clear_settings():
    p = _settings_path()
    if os.path.isfile(p):
        os.remove(p)


def _settings_has_ccc():
    p = _settings_path()
    if not os.path.isfile(p):
        return False
    with open(p, "r", encoding="utf-8") as f:
        s = json.load(f)
    return s.get("field_mapping", {}).get("ccc") == "客户"


def _clear_custom_key(key):
    fm.CUSTOM_HEADER_MAP.pop(key, None)


def _drain(worker):
    """取出 worker.queue 中的全部消息，返回消息列表。"""
    msgs = []
    while True:
        try:
            msgs.append(worker.queue.get_nowait())
        except Exception:
            break
    return msgs


def fresh_app():
    root = tk.Tk()
    root.withdraw()
    try:
        return root, SummaryApp(root)
    except Exception:
        root.destroy()
        raise


# 桩替换 messagebox，避免对话框在测试中弹窗阻塞
class _MsgStub:
    def __init__(self):
        self.calls = []

    def showwarning(self, *a, **k):
        self.calls.append(("warn", a, k))

    def showerror(self, *a, **k):
        self.calls.append(("err", a, k))

    def showinfo(self, *a, **k):
        self.calls.append(("info", a, k))


_orig_messagebox = gui.messagebox
gui.messagebox = _MsgStub()


def _set_dialog_mapping(dialog):
    """把对话框里的 ccc → 客户、金额 → 金额 设置好。"""
    for col_idx, header_text, var, combo in dialog._combobox_rows:
        if header_text == "ccc":
            var.set("客户")
        elif header_text == "金额":
            var.set("金额")


# ══════════════════════════════════════════════════════════════
#  T1: 正常表头自动识别
# ══════════════════════════════════════════════════════════════
print("\n[T1] 正常表头自动识别")
report = aggregate_with_report(FOLDER, logs_dir, ["a.xlsx"], {})
check("T1a: a.xlsx 成功处理", report["processed_count"] == 1,
      f"processed={report['processed_count']}")
check("T1b: 无跳过文件", report["skipped"] == [], f"skipped={report['skipped']}")
check("T1c: 识别到公司A", "公司A" in report["total"], f"total={list(report['total'])}")

# ══════════════════════════════════════════════════════════════
#  T2: 异常表头无法自动识别
# ══════════════════════════════════════════════════════════════
print("\n[T2] 异常表头「ccc」无法自动识别")
_clear_custom_key("ccc")
report = aggregate_with_report(FOLDER, logs_dir, ["b.xlsx"], {})
check("T2a: b.xlsx 被跳过", len(report["skipped"]) == 1,
      f"skipped={report['skipped']}")
if report["skipped"]:
    item = report["skipped"][0]
    check("T2b: 原因为 core_field", item.get("reason") == "core_field",
          f"reason={item.get('reason')}")
    check("T2c: 缺少字段为 客户", "客户" in item.get("missing", []),
          f"missing={item.get('missing')}")

# ══════════════════════════════════════════════════════════════
#  T3: 手动映射 + 不记住 → 当前任务仍成功
# ══════════════════════════════════════════════════════════════
print("\n[T3] 手动映射（不记住）仍成功")
_clear_custom_key("ccc")
report = aggregate_with_report(FOLDER, logs_dir, ["b.xlsx"], {"b.xlsx": {"ccc": "客户"}})
check("T3a: b.xlsx 成功处理", report["processed_count"] == 1,
      f"processed={report['processed_count']}")
check("T3b: 无跳过文件", report["skipped"] == [], f"skipped={report['skipped']}")
check("T3c: 识别到公司B", "公司B" in report["total"], f"total={list(report['total'])}")

# ══════════════════════════════════════════════════════════════
#  T4-T6: 组合输出（a 正常 + b 异常映射）
# ══════════════════════════════════════════════════════════════
print("\n[T4-T6] a.xlsx + b.xlsx 组合输出")
_clear_custom_key("ccc")
report = aggregate_with_report(
    FOLDER, logs_dir, ["a.xlsx", "b.xlsx"], {"b.xlsx": {"ccc": "客户"}},
)
check("T4a: 两个文件都成功处理", report["processed_count"] == 2,
      f"processed={report['processed_count']}")
check("T4b: 输出包含公司A", "公司A" in report["total"])
check("T4c: 输出包含公司B", "公司B" in report["total"])
check("T5: 客户数量 = 2", len(report["total"]) == 2, f"count={len(report['total'])}")
sum_money = sum(v.get("金额", 0) for v in report["total"].values())
check("T6: 总金额 = 300", abs(sum_money - 300) < 0.001, f"sum={sum_money}")

# ══════════════════════════════════════════════════════════════
#  T7: 可选字段（电话）缺失不丢失客户数据
# ══════════════════════════════════════════════════════════════
print("\n[T7] 可选字段缺失不丢失客户数据")
report = aggregate_with_report(FOLDER, logs_dir, ["a.xlsx", "e.xlsx"], {})
check("T7a: 两个文件都成功处理", report["processed_count"] == 2,
      f"processed={report['processed_count']}")
check("T7b: 公司A（含电话）保留", "公司A" in report["total"])
check("T7c: 公司E（缺电话）保留", "公司E" in report["total"])
check("T7d: 客户数量 = 2", len(report["total"]) == 2, f"count={len(report['total'])}")

# ══════════════════════════════════════════════════════════════
#  T8: 勾选「记住」写入 settings.json
# ══════════════════════════════════════════════════════════════
print("\n[T8] 勾选「记住」写入 settings.json")
_clear_settings()
_clear_custom_key("ccc")
root = tk.Tk()
dialog = gui.FieldMappingDialog(root, MAP_PATH, config_dir)
dialog.remember_var.set(True)
_set_dialog_mapping(dialog)
dialog._on_confirm()
check("T8a: 对话框 result=True", dialog.result is True)
check("T8b: 对话框 mappings 含 ccc→客户", dialog.mappings.get("ccc") == "客户",
      f"mappings={dialog.mappings}")
check("T8c: settings.json 写入 ccc→客户", _settings_has_ccc())
root.destroy()
_clear_custom_key("ccc")  # 清理，避免影响后续测试

# ══════════════════════════════════════════════════════════════
#  T9: 不勾选「记住」不写 settings.json
# ══════════════════════════════════════════════════════════════
print("\n[T9] 不勾选「记住」不写 settings.json")
_clear_settings()
_clear_custom_key("ccc")
root = tk.Tk()
dialog = gui.FieldMappingDialog(root, MAP_PATH, config_dir)
dialog.remember_var.set(False)
_set_dialog_mapping(dialog)
dialog._on_confirm()
check("T9a: 对话框 mappings 含 ccc→客户（临时生效）", dialog.mappings.get("ccc") == "客户",
      f"mappings={dialog.mappings}")
check("T9b: settings.json 不写 ccc", not _settings_has_ccc())
check("T9c: 内存映射也不含 ccc", "ccc" not in fm.CUSTOM_HEADER_MAP)
root.destroy()

# ══════════════════════════════════════════════════════════════
#  T10: 不记住 + 重启后临时映射消失
# ══════════════════════════════════════════════════════════════
print("\n[T10] 不记住 + 重启后临时映射消失")
_clear_settings()
_clear_custom_key("ccc")
root, app = fresh_app()
check("T10a: 新应用 task_field_mappings 为空", app.task_field_mappings == {},
      f"task_field_mappings={app.task_field_mappings}")
root.destroy()
# 重启后（无持久化映射），ccc 仍无法识别
report = aggregate_with_report(FOLDER, logs_dir, ["b.xlsx"], {})
check("T10b: 重启后 ccc 仍无法识别（被跳过）", len(report["skipped"]) == 1,
      f"skipped={report['skipped']}")

# ══════════════════════════════════════════════════════════════
#  T11: 多个文件中一个异常（按文件/表头上下文映射）
# ══════════════════════════════════════════════════════════════
print("\n[T11] 多个文件按文件上下文映射")
_clear_custom_key("ccc")
report = aggregate_with_report(
    FOLDER, logs_dir, ["a.xlsx", "b.xlsx", "c.xlsx"],
    {"b.xlsx": {"ccc": "客户"}, "c.xlsx": {"对手方": "客户"}},
)
check("T11a: 三个文件都成功处理", report["processed_count"] == 3,
      f"processed={report['processed_count']}")
check("T11b: 无跳过文件", report["skipped"] == [], f"skipped={report['skipped']}")
check("T11c: 三个公司都保留",
      all(c in report["total"] for c in ("公司A", "公司B", "公司C")),
      f"total={list(report['total'])}")

# ══════════════════════════════════════════════════════════════
#  T12: 核心字段无法确认 → 阻止生成汇总
# ══════════════════════════════════════════════════════════════
print("\n[T12] 核心字段无法确认 → 阻止生成汇总")
report = aggregate_with_report(FOLDER, logs_dir, ["d.xlsx"], {})
check("T12a: d.xlsx 被跳过", len(report["skipped"]) == 1,
      f"skipped={report['skipped']}")
if report["skipped"]:
    missing = report["skipped"][0].get("missing", [])
    check("T12b: 缺少 客户 和 金额", "客户" in missing and "金额" in missing,
          f"missing={missing}")

# WorkerThread 层验证：存在跳过文件时不生成输出文件
out_dir = os.path.join(TMP, "输出")
out_path = os.path.join(out_dir, "汇总.xlsx")
worker = gui._WorkerThread(FOLDER, out_path, logs_dir, file_names=["d.xlsx"], field_mappings={})
worker._pipeline()
msgs = _drain(worker)
err_msgs = [m for m in msgs if m[0] == "error"]
check("T12c: WorkerThread 发出 error 消息", len(err_msgs) == 1,
      f"msgs={[m[0] for m in msgs]}")
check("T12d: 未生成输出文件", not os.path.isfile(out_path),
      f"out_path={out_path}")

# ══════════════════════════════════════════════════════════════
#  T13: 错误信息清晰
# ══════════════════════════════════════════════════════════════
print("\n[T13] 错误信息清晰")
msg = gui._build_skip_error(report)
check("T13a: 含「汇总无法继续」", "汇总无法继续" in msg)
check("T13b: 含文件名 d.xlsx", "d.xlsx" in msg)
check("T13c: 含缺少字段提示", "缺少" in msg)
check("T13d: 含引导重新汇总", "重新汇总" in msg)

# ══════════════════════════════════════════════════════════════
#  T14: 可选字段缺失允许继续汇总
# ══════════════════════════════════════════════════════════════
print("\n[T14] 可选字段缺失允许继续汇总")
report = aggregate_with_report(FOLDER, logs_dir, ["e.xlsx"], {})
check("T14a: e.xlsx 成功处理", report["processed_count"] == 1,
      f"processed={report['processed_count']}")
check("T14b: 无跳过文件", report["skipped"] == [], f"skipped={report['skipped']}")
check("T14c: 公司E 保留", "公司E" in report["total"])

# ══════════════════════════════════════════════════════════════
#  T15: 取消不改变原映射
# ══════════════════════════════════════════════════════════════
print("\n[T15] 取消不改变原映射")
# 对话框层：取消 → result=False、mappings 空
root = tk.Tk()
dialog = gui.FieldMappingDialog(root, MAP_PATH, config_dir)
dialog._on_cancel()
check("T15a: 取消后 result=False", dialog.result is False)
check("T15b: 取消后 mappings 为空", dialog.mappings == {}, f"mappings={dialog.mappings}")
root.destroy()

# SummaryApp 层：对话框取消时，task_field_mappings 不变
class _FakeCancelDialog:
    def __init__(self, parent, file_path, base_dir):
        self.result = False
        self.mappings = {}


_orig_field_dialog = gui.FieldMappingDialog
gui.FieldMappingDialog = _FakeCancelDialog
root, app = fresh_app()
root.wait_window = lambda *a, **k: None
app.selected_folder.set(FOLDER)
app.selected_excel_files = ["b.xlsx"]
app._rescan_selected_files = lambda: None
app.task_field_mappings = {"b.xlsx": {"ccc": "客户"}}   # 已有映射
app._on_open_mapping_dialog("b.xlsx")
check("T15c: 取消后 task_field_mappings 不变",
      app.task_field_mappings.get("b.xlsx") == {"ccc": "客户"},
      f"task_field_mappings={app.task_field_mappings}")
root.destroy()
gui.FieldMappingDialog = _orig_field_dialog

# ══════════════════════════════════════════════════════════════
#  T16: 连续两个任务互不污染
# ══════════════════════════════════════════════════════════════
print("\n[T16] 连续两个任务互不污染")
root, app = fresh_app()
app.task_field_mappings = {"b.xlsx": {"ccc": "客户"}}   # 任务 1 的临时映射
app._on_start_new_task()                                # 开始任务 2
check("T16a: 新建任务后 task_field_mappings 清空", app.task_field_mappings == {},
      f"task_field_mappings={app.task_field_mappings}")
root.destroy()

# _on_folder_selected 也清空
root, app = fresh_app()
app.task_field_mappings = {"b.xlsx": {"ccc": "客户"}}
app.selected_folder.set(FOLDER)
app._clear_status = lambda: None
app._update_file_list_display = lambda: None
app.start_btn = type("B", (), {"config": lambda *a, **k: None})()
app._prompt_select_files = lambda folder, files: list(files)
app._rescan_selected_files = lambda: None
app._on_folder_selected()
check("T16b: 重选文件夹后 task_field_mappings 清空", app.task_field_mappings == {},
      f"task_field_mappings={app.task_field_mappings}")
root.destroy()

# ══════════════════════════════════════════════════════════════
#  版本统一
# ══════════════════════════════════════════════════════════════
print("\n[版本] 版本统一")
check("APP_VERSION == v1.5.4", APP_VERSION == "v1.5.4", f"实际 {APP_VERSION}")
check("APP_FULL_NAME 含 v1.5.4", "v1.5.4" in APP_FULL_NAME, f"实际 {APP_FULL_NAME}")

# ══════════════════════════════════════════════════════════════
#  清理
# ══════════════════════════════════════════════════════════════
utils.get_base_dir = _orig_utils_base_dir
gui.get_base_dir = _orig_gui_base_dir
gui.messagebox = _orig_messagebox
gui.FieldMappingDialog = _orig_field_dialog
fm.CUSTOM_HEADER_MAP.clear()
fm.CUSTOM_HEADER_MAP.update(_orig_custom_map)
reset_field_config(config_dir)
shutil.rmtree(TMP, ignore_errors=True)

print("\n" + "=" * 60)
print(f"  {n_passed}/{n_total} 通过")
print("=" * 60)
if n_passed != n_total:
    sys.exit(1)
