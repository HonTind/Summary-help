"""针对 4 个修复的验证测试（v1.5.0 适配）"""
import os, sys, shutil, tempfile

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

from openpyxl import Workbook
from Excel自动汇总助手.excel_reader import read_and_aggregate
from Excel自动汇总助手.processor import sort_by_amount
from Excel自动汇总助手.excel_writer import generate_summary

TEST_DIR = os.path.join(PROJECT_ROOT, "tests", "fix_verify")
if os.path.isdir(TEST_DIR):
    shutil.rmtree(TEST_DIR)
os.makedirs(TEST_DIR)

print("=" * 50)
print("  修复验证测试（v1.5.0）")
print("=" * 50)

# --- 修复3: data_only=True — 公式金额应被正确计算 ---
print("\n[修复3] data_only=True — 公式金额测试:")
wb = Workbook()
ws = wb.active
ws.append(["客户", "金额"])
ws.append(["公司A", 1000])
ws["B3"] = "=B2*2"   # 公式：2000
ws["A3"] = "公司B"
formula_path = os.path.join(TEST_DIR, "formula.xlsx")
wb.save(formula_path)
wb.close()

total, fc, field_names = read_and_aggregate(TEST_DIR, PROJECT_ROOT)
# v1.5.0: total values are dicts {field_name: value}
sum_a = total.get("公司A", {}).get("金额", 0)
sum_b = total.get("公司B", {}).get("金额", 0)
if sum_b == 2000:
    print(f"  PASS: 公式计算正确 — A={sum_a}, B={sum_b} (公式=B2*2)")
else:
    print(f"  FAIL: 公式未正确计算 — A={sum_a}, B={sum_b}, 期望 B=2000")

# --- 修复4: bool 类型拒绝 ---
print("\n[修复4] bool 类型拒绝测试:")
os.remove(formula_path)
wb2 = Workbook()
ws2 = wb2.active
ws2.append(["客户", "金额"])
ws2.append(["公司A", 1000.00])
ws2.append(["公司B", True])
ws2.append(["公司C", False])
ws2.append(["公司D", 500.00])
bool_path = os.path.join(TEST_DIR, "bool_test.xlsx")
wb2.save(bool_path)
wb2.close()

total2, fc2, field_names2 = read_and_aggregate(TEST_DIR, PROJECT_ROOT)
# v1.5.0: total values are dicts, sum the money field
expected_sum = 1500.00
actual_sum = sum(item.get("金额", 0) for item in total2.values())
if "公司B" not in total2 and "公司C" not in total2:
    if abs(actual_sum - expected_sum) < 0.01:
        print(f"  PASS: bool 值被正确跳过 — customers={list(total2.keys())}, sum={actual_sum}")
    else:
        print(f"  FAIL: 汇总金额不对 — sum={actual_sum}, 期望={expected_sum}")
else:
    print(f"  FAIL: bool 值未被跳过 — customers={list(total2.keys())}")

# --- 修复1&2: GUI save错误日志 (手动检查) ---
print("\n[修复1&2] GUI文件占用+错误日志:")
print("  手动验证步骤:")
print("    1. 启动 GUI: python -m Excel自动汇总助手.gui")
print("    2. 选一个包含正常Excel的文件夹")
print("    3. 打开汇总.xlsx（用Excel打开）")
print("    4. 点'开始汇总'")
print("    5. 应弹出: '汇总文件正在打开，请关闭后重新生成。'")
print("    6. 检查 error.log 是否记录了保存失败")

print("\n" + "=" * 50)
print("  验证完成")
print("=" * 50)

shutil.rmtree(TEST_DIR)
