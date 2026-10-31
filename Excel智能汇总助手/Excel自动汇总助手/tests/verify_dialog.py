"""
v1.3 第二阶段 — FieldMappingDialog 验证测试

测试:
  T1. 对话框打开 — 表头正确加载
  T2. 预选 — diagnose_failure 部分匹配 → 下拉框预填
  T3. 验证 — 客户和金额都未选 → 警告
  T4. 验证 — 金额未选 → 警告
  T5. 验证 — 客户未选 → 警告
  T6. 保存 — 勾选记住 → settings.json 写入
  T7. 保存 — 不勾选记住 → settings.json 不写入
  T8. 取消 — result=False
  T9. 确定 — 全部合法 → result=True

运行:
  cd "Day07(小试牛刀)"
  python -m Excel自动汇总助手.tests.verify_dialog
"""

import os, sys, json, shutil, time

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

import tkinter as tk
from openpyxl import Workbook

import Excel自动汇总助手.config as config_module
from Excel自动汇总助手.field_mapper import (
    load_settings, save_field_mapping, remove_field_mapping, get_all_mappings,
)

TEST_DIR = os.path.join(PROJECT_ROOT, "tests", "_dialog_test")
if os.path.isdir(TEST_DIR):
    shutil.rmtree(TEST_DIR)
os.makedirs(TEST_DIR)

_original_map = dict(config_module.CUSTOM_HEADER_MAP)

def reset_map():
    config_module.CUSTOM_HEADER_MAP.clear()
    config_module.CUSTOM_HEADER_MAP.update(_original_map)

passed = 0
total = 0

def check(name, condition, detail=""):
    global passed, total
    total += 1
    if condition:
        passed += 1
        print(f"  PASS  {name}")
    else:
        print(f"  FAIL  {name}  {detail}")

print("=" * 55)
print("  v1.3 FieldMappingDialog 验证")
print("=" * 55)

# ═══════════════════════════════════════════════════════
#  准备：创建测试 Excel 文件
# ═══════════════════════════════════════════════════════

# 文件A：关键词部分匹配（客户能识别，金额不能）
test_file_a = os.path.join(TEST_DIR, "test_a.xlsx")
wb = Workbook()
ws = wb.active
ws.append(["客户", "价格合计", "日期", "经手人"])
ws.append(["公司A", 1000, "2024-01-01", "张三"])
wb.save(test_file_a)
wb.close()

# 文件B：完全不能识别
test_file_b = os.path.join(TEST_DIR, "test_b.xlsx")
wb2 = Workbook()
ws2 = wb2.active
ws2.append(["甲方名称", "合同价款", "签约日", "甲方代表"])
ws2.append(["公司B", 2000, "2024-02-01", "李四"])
wb2.save(test_file_b)
wb2.close()

# ═══════════════════════════════════════════════════════
#  创建隐藏的 tk root（对话框需要 parent）
# ═══════════════════════════════════════════════════════
root = tk.Tk()
root.withdraw()  # 不显示主窗口

from Excel自动汇总助手.gui import FieldMappingDialog

# ═══════════════════════════════════════════════════════
#  T1: 打开对话框 — 表头正确加载
# ═══════════════════════════════════════════════════════
print("\n[T1] 打开对话框 — 表头正确加载")
reset_map()

# 不显示（headless 测试），仅检查内部状态
dialog = FieldMappingDialog(root, test_file_a, TEST_DIR)
expected_headers = [(0, "客户"), (1, "价格合计"), (2, "日期"), (3, "经手人")]
actual_headers = dialog._header_rows
check("表头行数正确", len(actual_headers) == 4,
      f"期望 4, 实际 {len(actual_headers)}")
check("表头内容正确", actual_headers == expected_headers,
      f"\n      期望 {expected_headers}\n      实际 {actual_headers}")
dialog.destroy()

# ═══════════════════════════════════════════════════════
#  T2: 预选 — 部分匹配
# ═══════════════════════════════════════════════════════
print("\n[T2] 预选 — diagnose_failure 部分匹配")

dialog = FieldMappingDialog(root, test_file_a, TEST_DIR)
pre = dialog._preselect()
# v1.5.0: 字段名直接使用配置中的名称（中文），不再使用 "customer"/"money" 内部码
check("'客户' 被预选为 客户", pre.get(0) == "客户",
      f"实际: {pre}")
# "价格合计" 不在金额关键词中 → 不应预选
check("'价格合计' 未被预选", 1 not in pre,
      f"实际: {pre}")
dialog.destroy()

# 文件B：完全不能识别 → preselected 应为空
dialog2 = FieldMappingDialog(root, test_file_b, TEST_DIR)
pre2 = dialog2._preselect()
check("完全未知表头 — 预选为空", pre2 == {},
      f"实际: {pre2}")
dialog2.destroy()

# ═══════════════════════════════════════════════════════
#  T3: 验证 — 客户和金额都未选 → 警告
# ═══════════════════════════════════════════════════════
print("\n[T3] 验证 — 客户和金额都未选")

dialog = FieldMappingDialog(root, test_file_b, TEST_DIR)
# 把所有下拉框设为 "忽略此列"
for _, _, var, _combo in dialog._combobox_rows:
    var.set("忽略此列")

# 调用确定（不显示 dialog，直接测试逻辑）
triggered = {"called": False}
_orig_warn = tk.messagebox.showwarning
def fake_warn(title, msg, **kw):
    triggered["title"] = title
    triggered["msg"] = msg
    triggered["called"] = True
    return "ok"
tk.messagebox.showwarning = fake_warn
dialog._on_confirm()
tk.messagebox.showwarning = _orig_warn

check("弹出警告", triggered["called"])
check("消息包含'必要的信息'", "必要的信息" in triggered.get("msg", ""),
      f"实际: {triggered.get('msg','')}")
check("对话框未关闭（未 destroy）", dialog.winfo_exists())
dialog.destroy()

# ═══════════════════════════════════════════════════════
#  T4: 验证 — 金额未选
# ═══════════════════════════════════════════════════════
print("\n[T4] 验证 — 客户已选，金额未选")

dialog = FieldMappingDialog(root, test_file_b, TEST_DIR)
for col_idx, header_text, var, _combo in dialog._combobox_rows:
    if "甲方名称" in header_text:
        var.set("客户")
    else:
        var.set("忽略此列")

triggered = {"called": False}
def fake_warn2(title, msg, **kw):
    triggered["title"] = title
    triggered["msg"] = msg
    triggered["called"] = True
    return "ok"
tk.messagebox.showwarning = fake_warn2
dialog._on_confirm()
tk.messagebox.showwarning = _orig_warn

check("弹出警告", triggered["called"])
# v1.5.0: 消息使用字段名「金额」
check("消息包含'金额'", "金额" in triggered.get("msg", ""),
      f"实际: {triggered.get('msg','')}")
dialog.destroy()

# ═══════════════════════════════════════════════════════
#  T5: 验证 — 客户未选
# ═══════════════════════════════════════════════════════
print("\n[T5] 验证 — 金额已选，客户未选")

dialog = FieldMappingDialog(root, test_file_b, TEST_DIR)
for col_idx, header_text, var, _combo in dialog._combobox_rows:
    if "合同价款" in header_text:
        var.set("金额")
    else:
        var.set("忽略此列")

triggered = {"called": False}
def fake_warn3(title, msg, **kw):
    triggered["title"] = title
    triggered["msg"] = msg
    triggered["called"] = True
    return "ok"
tk.messagebox.showwarning = fake_warn3
dialog._on_confirm()
tk.messagebox.showwarning = _orig_warn

check("弹出警告", triggered["called"])
# v1.5.0: 消息使用字段名「客户」
check("消息包含'客户'", "客户" in triggered.get("msg", ""),
      f"实际: {triggered.get('msg','')}")
dialog.destroy()

# ═══════════════════════════════════════════════════════
#  T6: 保存 — 勾选记住 → settings.json 写入
# ═══════════════════════════════════════════════════════
print("\n[T6] 保存 — 勾选记住 → settings.json 写入")

reset_map()
# 确保 settings.json 不存在或为空
settings_path = os.path.join(TEST_DIR, "settings.json")
if os.path.isfile(settings_path):
    os.remove(settings_path)

dialog = FieldMappingDialog(root, test_file_b, TEST_DIR)
# 设置选择
for col_idx, header_text, var, _combo in dialog._combobox_rows:
    if "甲方名称" in header_text:
        var.set("客户")
    elif "合同价款" in header_text:
        var.set("金额")
    elif "签约日" in header_text:
        var.set("客户")  # v1.5.0: "日期"不在默认配置中，用"客户"代替
    else:
        var.set("忽略此列")

dialog.remember_var.set(True)
dialog._on_confirm()

# 检查 settings.json
check("settings.json 已创建", os.path.isfile(settings_path))
if os.path.isfile(settings_path):
    with open(settings_path, "r", encoding="utf-8") as f:
        s = json.load(f)
    fm = s.get("field_mapping", {})
    # v1.5.0: 映射值使用字段名（中文），不再使用 "customer"/"money" 内部码
    check("甲方名称 → 客户", fm.get("甲方名称") == "客户",
          f"实际: {fm}")
    check("合同价款 → 金额", fm.get("合同价款") == "金额",
          f"实际: {fm}")
    # v1.5.0: "日期" 不在默认字段配置中，此处不测试
    check("甲方代表 不在 JSON（ignore 不保存）",
          "甲方代表" not in fm,
          f"实际: {fm}")

# 检查内存映射（v1.5.0: 使用字段名）
mem = get_all_mappings()
check("内存映射: 甲方名称→客户",
      mem.get("甲方名称") == "客户",
      f"实际: {mem}")
check("内存映射: 合同价款→金额",
      mem.get("合同价款") == "金额",
      f"实际: {mem}")

# ═══════════════════════════════════════════════════════
#  T7: 保存 — 不勾选记住 → 不写 settings.json
# ═══════════════════════════════════════════════════════
print("\n[T7] 不勾选记住 → settings.json 不变化")

# 记录当前 JSON 内容
before_json = ""
if os.path.isfile(settings_path):
    with open(settings_path, "r", encoding="utf-8") as f:
        before_json = f.read()

dialog = FieldMappingDialog(root, test_file_a, TEST_DIR)
for col_idx, header_text, var, _combo in dialog._combobox_rows:
    if "客户" in header_text:
        var.set("客户")
    elif "价格" in header_text:
        var.set("金额")
    else:
        var.set("忽略此列")

dialog.remember_var.set(False)
dialog._on_confirm()

after_json = ""
if os.path.isfile(settings_path):
    with open(settings_path, "r", encoding="utf-8") as f:
        after_json = f.read()

check("settings.json 未变化", before_json == after_json,
      f"变化了！\n      before={before_json}\n      after={after_json}")

# ═══════════════════════════════════════════════════════
#  T8: 取消 — result=False
# ═══════════════════════════════════════════════════════
print("\n[T8] 取消 — result=False")

dialog = FieldMappingDialog(root, test_file_a, TEST_DIR)
dialog._on_cancel()
check("result=False", dialog.result is False)
check("窗口已销毁", not dialog.winfo_exists())

# ═══════════════════════════════════════════════════════
#  T9: 确定 — 全部合法 → result=True
# ═══════════════════════════════════════════════════════
print("\n[T9] 确定 — result=True")

dialog = FieldMappingDialog(root, test_file_a, TEST_DIR)
for col_idx, header_text, var, _combo in dialog._combobox_rows:
    if "客户" in header_text:
        var.set("客户")
    elif "价格" in header_text:
        var.set("金额")
    else:
        var.set("忽略此列")
dialog.remember_var.set(False)
dialog._on_confirm()
check("result=True", dialog.result is True)
check("窗口已销毁", not dialog.winfo_exists())

# ═══════════════════════════════════════════════════════
#  清理
# ═══════════════════════════════════════════════════════
root.destroy()
reset_map()
shutil.rmtree(TEST_DIR)

print(f"\n{'=' * 55}")
print(f"  {passed}/{total} 通过")
print(f"{'=' * 55}")
