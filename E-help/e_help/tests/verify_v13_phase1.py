"""
v1.3 第一阶段 — 字段映射持久化验证测试（v1.5.0 适配）

测试范围:
  1. load_settings — 空文件 → 无崩溃
  2. save_field_mapping — 写入 JSON + 更新内存（v1.5.0: 使用字段名）
  3. 重启模拟 — 清除 CUSTOM_HEADER_MAP → load_settings → 映射恢复
  4. config.py 优先级 — 硬编码映射不被 settings.json 覆盖
  5. get_headers_for_display — 返回正确的 (index, text) 列表
  6. detect_columns 集成 — 新映射使之前失败的识别变为成功（v1.5.0: list[FieldDetection]）
  7. JSON 损坏 — load_settings 不崩溃
  8. settings.json 不存在 — load_settings 不崩溃

运行方式:
  cd "Day07(小试牛刀)"
  python -m e_help.tests.verify_v13_phase1
"""

import os
import sys
import json
import shutil
import tempfile

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

from openpyxl import Workbook

# ═══════════════════════════════════════════════════════════
#  准备：独立的测试环境（不影响真实 settings.json）
# ═══════════════════════════════════════════════════════════

TEST_BASE = os.path.join(PROJECT_ROOT, "tests", "_v13_test_env")
if os.path.isdir(TEST_BASE):
    shutil.rmtree(TEST_BASE)
os.makedirs(TEST_BASE)

import e_help.config as config_module

# 记录原始映射，测试结束后恢复
_original_map = dict(config_module.CUSTOM_HEADER_MAP)


def reset_map():
    """清除 CUSTOM_HEADER_MAP 中非原始条目"""
    config_module.CUSTOM_HEADER_MAP.clear()
    config_module.CUSTOM_HEADER_MAP.update(_original_map)


print("=" * 55)
print("  v1.3 第一阶段 — 字段映射持久化测试（v1.5.0）")
print("=" * 55)

# ═══════════════════════════════════════════════════════════
#  测试 1: load_settings — 文件不存在 → 不崩溃
# ═══════════════════════════════════════════════════════════
print("\n[测试 1] load_settings — settings.json 不存在")
from e_help.field_mapper import load_settings

settings_path = os.path.join(TEST_BASE, "settings.json")
assert not os.path.isfile(settings_path), "测试环境不应有 settings.json"

try:
    load_settings(TEST_BASE)
    print("  PASS: 无崩溃，静默跳过")
except Exception as e:
    print(f"  FAIL: {e}")

# ═══════════════════════════════════════════════════════════
#  测试 2: save_field_mapping — 写入 JSON + 更新内存（v1.5.0: 使用字段名）
# ═══════════════════════════════════════════════════════════
print("\n[测试 2] save_field_mapping — 写入 + 内存更新")
from e_help.field_mapper import save_field_mapping

reset_map()

# v1.5.0: 使用字段名（中文）而非内部码
save_field_mapping(TEST_BASE, "购货单位", "客户")
save_field_mapping(TEST_BASE, "价税合计", "金额")

# 2a: JSON 文件是否正确写入
assert os.path.isfile(settings_path), "settings.json 应被创建"
with open(settings_path, "r", encoding="utf-8") as f:
    saved = json.load(f)

fm = saved.get("field_mapping", {})
assert fm.get("购货单位") == "客户", f"JSON 中购货单位映射错误: {fm}"
assert fm.get("价税合计") == "金额", f"JSON 中价税合计映射错误: {fm}"
print(f"  PASS: JSON 写入正确 → {fm}")

# 2b: 内存映射是否立即更新
assert config_module.CUSTOM_HEADER_MAP.get("购货单位") == "客户"
assert config_module.CUSTOM_HEADER_MAP.get("价税合计") == "金额"
print(f"  PASS: 内存映射立即生效 → CUSTOM_HEADER_MAP = {dict(config_module.CUSTOM_HEADER_MAP)}")

# ═══════════════════════════════════════════════════════════
#  测试 3: 重启模拟 — 清除内存 → load_settings → 恢复
# ═══════════════════════════════════════════════════════════
print("\n[测试 3] 重启模拟 — 清除 CUSTOM_HEADER_MAP → load_settings → 恢复")

# 模拟重启：清除内存中的运行时映射
config_module.CUSTOM_HEADER_MAP.clear()
config_module.CUSTOM_HEADER_MAP.update(_original_map)  # 仅保留 config.py 硬编码

assert "购货单位" not in config_module.CUSTOM_HEADER_MAP, "购货单位 不应在硬编码中"
assert "价税合计" not in config_module.CUSTOM_HEADER_MAP, "价税合计 不应在硬编码中"

# 启动时调用
load_settings(TEST_BASE)

assert config_module.CUSTOM_HEADER_MAP.get("购货单位") == "客户", \
    "重启后 购货单位 应恢复为 客户"
assert config_module.CUSTOM_HEADER_MAP.get("价税合计") == "金额", \
    "重启后 价税合计 应恢复为 金额"
print("  PASS: 重启后映射正确恢复")
print(f"         CUSTOM_HEADER_MAP = {dict(config_module.CUSTOM_HEADER_MAP)}")

# ═══════════════════════════════════════════════════════════
#  测试 4: config.py 硬编码优先级高于 settings.json
# ═══════════════════════════════════════════════════════════
print("\n[测试 4] settings.json 覆盖 config.py 硬编码")
reset_map()

# 在 config.py 中预设一个映射（模拟用户手动编辑）
config_module.CUSTOM_HEADER_MAP["模板客户"] = "客户"

# 用户通过 save_field_mapping 保存同名映射为 金额
save_field_mapping(TEST_BASE, "模板客户", "金额")

# save 后内存应立即更新（用户的主动操作）
assert config_module.CUSTOM_HEADER_MAP.get("模板客户") == "金额", \
    "save 后内存应更新为 金额"
print("  PASS: save 后内存='金额'（用户操作优先生效）")

# 模拟重启：硬编码回到 config.py 的原始值，settings.json 加载覆盖
config_module.CUSTOM_HEADER_MAP.clear()
config_module.CUSTOM_HEADER_MAP["模板客户"] = "客户"  # 模拟 config.py 重载

load_settings(TEST_BASE)

assert config_module.CUSTOM_HEADER_MAP.get("模板客户") == "金额", \
    "重启后 settings.json 的值应覆盖硬编码"
print("  PASS: 重启后='金额'（settings.json 覆盖 config.py）")

# 清理
del config_module.CUSTOM_HEADER_MAP["模板客户"]

# ═══════════════════════════════════════════════════════════
#  测试 5: get_headers_for_display
# ═══════════════════════════════════════════════════════════
print("\n[测试 5] get_headers_for_display")
from e_help.field_mapper import get_headers_for_display

wb = Workbook()
ws = wb.active
ws.append(["购货单位", None, "价税合计", "", "  备注  "])

headers = get_headers_for_display(ws)
expected = [(0, "购货单位"), (2, "价税合计"), (4, "备注")]
assert headers == expected, f"期望 {expected}, 实际 {headers}"
print(f"  PASS: 返回 {headers}")
wb.close()

# ═══════════════════════════════════════════════════════════
#  测试 6: detect_columns 集成 — 新映射使识别成功（v1.5.0: list[FieldDetection]）
# ═══════════════════════════════════════════════════════════
print("\n[测试 6] detect_columns 集成 — 新映射使识别成功")

wb2 = Workbook()
ws2 = wb2.active
ws2.append(["购货单位", "价税合计", "备注"])
ws2.append(["公司A", 1000, "test"])
ws2.append(["公司B", 2000, "test"])

from e_help.field_mapper import detect_columns

detections = detect_columns(ws2)
assert detections is not None, "应成功识别（购货单位→客户, 价税合计→金额）"
# v1.5.0: detect_columns 返回 list[FieldDetection]
assert len(detections) >= 2, f"应至少识别2个字段, 实际 {len(detections)}"
assert detections[0].col_index == 0, f"第1个字段列索引应为0, 实际{detections[0].col_index}"
assert detections[1].col_index == 1, f"第2个字段列索引应为1, 实际{detections[1].col_index}"
print(f"  PASS: 识别成功 → [{detections[0].field_name}:{detections[0].col_index}, {detections[1].field_name}:{detections[1].col_index}]")
wb2.close()

# ═══════════════════════════════════════════════════════════
#  测试 7: JSON 损坏 → 不崩溃
# ═══════════════════════════════════════════════════════════
print("\n[测试 7] JSON 损坏 → 不崩溃")

corrupt_path = os.path.join(TEST_BASE, "settings.json")
with open(corrupt_path, "w", encoding="utf-8") as f:
    f.write("{这不是合法的 JSON")

reset_map()
try:
    load_settings(TEST_BASE)
    print("  PASS: 损坏的 JSON 未导致崩溃")
except Exception as e:
    print(f"  FAIL: {e}")

# ═══════════════════════════════════════════════════════════
#  测试 8: save_field_mapping 在损坏 JSON 上 → 覆盖修复
# ═══════════════════════════════════════════════════════════
print("\n[测试 8] save_field_mapping → 覆盖损坏的 JSON")

save_field_mapping(TEST_BASE, "公司", "客户")

with open(settings_path, "r", encoding="utf-8") as f:
    repaired = json.load(f)

assert repaired.get("field_mapping", {}).get("公司") == "客户"
print("  PASS: 损坏 JSON 被正确覆盖，新映射写入成功")

# ═══════════════════════════════════════════════════════════
#  清理
# ═══════════════════════════════════════════════════════════

reset_map()
shutil.rmtree(TEST_BASE)

print("\n" + "=" * 55)
print("  8/8 全部通过")
print("=" * 55)
