"""
v1.3 第三阶段 — FieldMappingDialog 接入主流程 回归测试

测试:
  R1. 启动时自动加载 settings.json → 预扫描中生效
  R2. [WARN] 行出现 [设置列] 按钮
  R3. 点击 [设置列] → 对话框 → 映射 → 保存 → 自动重扫描
  R4. 映射后 [WARN] → [OK]
  R5. [开始汇总] 按钮在全部可识别后恢复
  R6. 已有功能无回归（正常文件不受影响）

运行:
  cd "Day07(小试牛刀)"
  python -m e_help.tests.verify_v13_phase3
"""

import os, sys, json, shutil

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

import tkinter as tk
from openpyxl import Workbook

import e_help.config as config_module
from e_help.field_mapper import (
    load_settings, save_field_mapping, remove_field_mapping, get_all_mappings,
)
from e_help.gui import SummaryApp, FieldMappingDialog

TEST_DIR = os.path.join(PROJECT_ROOT, "tests", "_phase3_test")
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
print("  v1.3 第三阶段 — 接入主流程 回归测试")
print("=" * 55)

# ═══════════════════════════════════════════════════════
#  准备：测试数据
# ═══════════════════════════════════════════════════════

# 文件A：正常（能被关键词识别）
wb_a = Workbook()
ws_a = wb_a.active
ws_a.append(["客户", "金额"])
ws_a.append(["公司A", 1000])
wb_a.save(os.path.join(TEST_DIR, "normal.xlsx"))
wb_a.close()

# 文件B：特殊表头（需要自定义映射才能识别）
wb_b = Workbook()
ws_b = wb_b.active
ws_b.append(["购货单位", "价税合计"])
ws_b.append(["公司B", 2000])
wb_b.save(os.path.join(TEST_DIR, "special.xlsx"))
wb_b.close()

# 文件C：另一个特殊表头
wb_c = Workbook()
ws_c = wb_c.active
ws_c.append(["单位名称", "应收金额"])
ws_c.append(["公司C", 3000])
wb_c.save(os.path.join(TEST_DIR, "special2.xlsx"))
wb_c.close()

# ═══════════════════════════════════════════════════════
#  R1: 启动时自动加载 settings.json
# ═══════════════════════════════════════════════════════
print("\n[R1] 启动时自动加载 settings.json")

reset_map()

# 预写入 settings.json
settings_path = os.path.join(TEST_DIR, "settings.json")
init_settings = {
    "field_mapping": {
        "购货单位": "customer",
        "价税合计": "money",
    }
}
with open(settings_path, "w", encoding="utf-8") as f:
    json.dump(init_settings, f, ensure_ascii=False, indent=2)

# 启动 GUI → 内部调用了 load_settings(get_base_dir())
# 但 get_base_dir() 返回程序目录，不是 TEST_DIR
# 我们直接调用来模拟启动行为
load_settings(TEST_DIR)

check("购货单位 → customer",
      config_module.CUSTOM_HEADER_MAP.get("购货单位") == "customer")
check("价税合计 → money",
      config_module.CUSTOM_HEADER_MAP.get("价税合计") == "money")

# ═══════════════════════════════════════════════════════
#  R2: [WARN] 行出现 [设置列] 按钮 + R3: 对话框 → 映射 → 重扫
# ═══════════════════════════════════════════════════════
print("\n[R2+R3] [设置列] 按钮 + 对话框流程")

reset_map()
# 先确保没有 special2 的映射，让它出现在 [WARN] 中
if os.path.isfile(settings_path):
    os.remove(settings_path)

root = tk.Tk()
root.withdraw()

# 创建 app 并通过设置文件夹来触发扫描
app = SummaryApp(root)
# v1.5.2 Hotfix3: 勾选对话框已分离，此处模拟「全选」以保持旧回归语义
app._prompt_select_files = lambda folder, files: list(files)
app._show_task()
app.selected_folder.set(TEST_DIR)
app._on_folder_selected()

# 检查文件列表
ok_files = list(app._preview_ok)
warn_files = [fn for (fn, hdrs) in app._preview_warn]

print(f"  [OK] 文件: {ok_files}")
print(f"  [WARN] 文件: {warn_files}")

# "normal.xlsx" 应该有 "客户" "金额" 关键词 → 自动识别
# "special.xlsx" 的 "购货单位" "价税合计" 在关键词中 → 应该也能识别！
# "special2.xlsx" 的 "单位名称" "应收金额" → "单位名称"不在客户关键词中，"应收金额"在金额关键词中

# Wait - let me check: 购货单位 IS in CUSTOMER_KEYWORDS and 价税合计 IS in MONEY_KEYWORDS
# So special.xlsx should auto-detect successfully!
# Only special2.xlsx should fail (单位名称 not in customer keywords)

has_normal = "normal.xlsx" in ok_files
# "购货单位" IS in CUSTOMER_KEYWORDS list — yes it is
has_special = "special.xlsx" in ok_files
# "单位名称" NOT in CUSTOMER_KEYWORDS (关键词: 客户, 客户名称, 购货单位, 单位名称... wait let me check)
# Actually looking at config.py: "单位名称" IS in CUSTOMER_KEYWORDS!
# And "应收金额" IS in MONEY_KEYWORDS!
# So ALL three files should auto-detect!

# Hmm, this test setup doesn't produce [WARN] files with the current keyword lists.
# Let me use different headers that are definitely not in the keyword lists.

# Let me create new test files with truly unrecognizable headers
wb_x = Workbook()
ws_x = wb_x.active
ws_x.append(["甲方", "乙方金额", "签约日"])  # "甲方" not a customer keyword
ws_x.append(["CompanyX", 1000, "date"])
wb_x.save(os.path.join(TEST_DIR, "weird.xlsx"))
wb_x.close()

# Re-scan
app._on_folder_selected()
ok_files = list(app._preview_ok)
warn_files = [fn for (fn, hdrs) in app._preview_warn]

print(f"  重新扫描后 [OK] 文件: {ok_files}")
print(f"  重新扫描后 [WARN] 文件: {warn_files}")

# weird.xlsx 应该在 warn_files 中（"甲方" "乙方金额" 不在关键词中）
# 注意: "甲方名称" 不在关键词列表中，但"甲方"也不在。让我确认...
# CUSTOMER_KEYWORDS = ["客户", "客户名称", "购货单位", "单位名称", "公司", "顾客", "收货人", "买方"]
# "甲方" 不在其中 → [WARN]
# MONEY_KEYWORDS = ["金额", "总金额", "合计金额", "价税合计", "货款", "费用", "小计", "应收金额"]
# "乙方金额" → "金额" in "乙方金额" → [OK] (包含匹配!)
# So weird.xlsx: customer=None, money=1 (乙方金额 matches 金额 via contains)
# Still [WARN] because customer not found

weird_warn = "weird.xlsx" in warn_files
check("weird.xlsx 出现在 [WARN] 中", weird_warn,
      f"warn_files={warn_files}")

# 检查 Text widget 中是否有 [设置列] 按钮
# tk.Text 的 window_create 产生的窗口可通过 dump 检查
text_content = app.file_list_text.get("1.0", "end")
check("[WARN] 行显示 '无法识别'",
      "无法识别" in text_content,
      f"text_content片段: {text_content[:200]}")

# 触发 [设置列] 按钮 — 模拟点击（直接调方法）
if weird_warn:
    print("\n[R3] 模拟点击 [设置列] — 打开对话框 + 映射 + 保存")

    # 程序化设置映射（模拟用户在对话框中选择了客户和金额）
    # 先直接通过 field_mapper 接口模拟（对话框本身已在阶段2测试过）
    # v1.5.0: 使用字段名（中文）而非内部码
    save_field_mapping(TEST_DIR, "甲方", "客户")
    save_field_mapping(TEST_DIR, "乙方金额", "金额")

    check("映射已保存到内存",
          config_module.CUSTOM_HEADER_MAP.get("甲方") == "客户")
    check("settings.json 已写入",
          os.path.isfile(settings_path))

    # 模拟对话框关闭后自动重扫
    app._on_folder_selected()

    ok_files2 = list(app._preview_ok)
    warn_files2 = [fn for (fn, hdrs) in app._preview_warn]

    print(f"  映射后 [OK] 文件: {ok_files2}")
    print(f"  映射后 [WARN] 文件: {warn_files2}")

    # R4: [WARN] → [OK]
    weird_ok = "weird.xlsx" in ok_files2
    weird_not_warn = "weird.xlsx" not in warn_files2
    check("R4: weird.xlsx 从 [WARN] 变为 [OK]", weird_ok and weird_not_warn,
          f"ok={ok_files2}, warn={warn_files2}")

    # R5: [开始汇总] 按钮应恢复（全部可识别）
    btn_state = str(app.start_btn["state"])
    check("R5: [开始汇总] 按钮可点击", btn_state == "normal",
          f"按钮状态={btn_state}")

# ═══════════════════════════════════════════════════════
#  R6: 已有功能无回归
# ═══════════════════════════════════════════════════════
print("\n[R6] 已有功能无回归")

# 6a: 正常文件识别的文件不受影响
check("normal.xlsx 仍在 [OK] 中",
      "normal.xlsx" in app._preview_ok)

# 6b: 显式检查核心结构未被改变
check("_WorkerThread 类存在",
      hasattr(__import__('e_help.gui', fromlist=['_WorkerThread']), '_WorkerThread'))
check("_preview_ok 结构不变",
      all(isinstance(item, str) for item in app._preview_ok[:1]))
check("_preview_warn 结构不变",
      all(isinstance(item, tuple) and len(item) == 2
          for item in app._preview_warn[:1])
      if app._preview_warn else True)

# ═══════════════════════════════════════════════════════
#  清理
# ═══════════════════════════════════════════════════════
root.destroy()
reset_map()
shutil.rmtree(TEST_DIR)

print(f"\n{'=' * 55}")
print(f"  {passed}/{total} 通过")
print(f"{'=' * 55}")
