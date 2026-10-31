"""
Excel自动汇总助手 v1.2.2.1 — 稳定性测试套件

测试范围：
  一、Excel异常输入测试（9项）
  二、文件异常测试（5项）
  三、大数据测试（性能）
  四、PyInstaller测试（手动检查清单）

运行方式：
  cd "Excel自动汇总助手"
  python tests/run_stability_tests.py

输出：
  - 终端彩色测试矩阵
  - tests/test_report.txt 详细报告
"""

import os
import sys
import time
import shutil
import tempfile
import traceback
import io
from contextlib import redirect_stdout
from datetime import datetime

# 项目根目录
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 父目录（Day07(小试牛刀)）加入 sys.path，使 "Excel自动汇总助手" 作为包可导入
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

from openpyxl import Workbook, load_workbook

# 导入核心模块（包内相对导入）
from Excel自动汇总助手.config import FOLDER_NAME, OUTPUT_FILE, ERROR_LOG
from Excel自动汇总助手.utils import get_base_dir
from Excel自动汇总助手.excel_reader import read_and_aggregate
from Excel自动汇总助手.processor import sort_by_amount
from Excel自动汇总助手.excel_writer import generate_summary
from Excel自动汇总助手.field_mapper import detect_columns, get_header_list, diagnose_failure

# ============================================================
#  测试工具
# ============================================================

TEST_DATA = os.path.join(PROJECT_ROOT, "tests", "test_data")
RESULTS = []  # [(test_id, category, sub_category, passed, detail, crash_traceback), ...]


def reset_test_data():
    """清空并重建测试数据目录"""
    if os.path.isdir(TEST_DATA):
        shutil.rmtree(TEST_DATA)
    os.makedirs(TEST_DATA)


def make_workbook(filename, headers, rows):
    """快速创建一个 xlsx 文件"""
    wb = Workbook()
    ws = wb.active
    if headers:
        ws.append(headers)
    for row in rows:
        ws.append(row)
    wb.save(filename)
    wb.close()


def record(test_id, category, sub_category, passed, detail="", crash_tb=None):
    """记录一条测试结果"""
    RESULTS.append((test_id, category, sub_category, passed, detail, crash_tb))
    status = "  PASS" if passed else "  FAIL"
    print(f"  [{status}] {test_id} : {sub_category}")
    if detail:
        print(f"         {detail}")
    if crash_tb:
        # 只打印最后 3 行
        lines = crash_tb.strip().split("\n")
        for line in lines[-4:]:
            print(f"         {line}")


def run_pipeline(folder):
    """运行完整汇总管线，返回 (成功, 异常信息)"""
    try:
        total, file_count = read_and_aggregate(folder, PROJECT_ROOT)
        if file_count == 0 or len(total) == 0:
            return True, f"file_count={file_count}, customers={len(total)} (无有效数据)"
        sorted_data = sort_by_amount(total)
        output_path = os.path.join(TEST_DATA, OUTPUT_FILE)
        generate_summary(sorted_data, total, output_path)
        return True, f"OK: {file_count} files, {len(total)} customers, sum={sum(total.values()):.2f}"
    except Exception as e:
        return False, f"{type(e).__name__}: {e}\n{traceback.format_exc()}"


def capture_read_only(folder):
    """只读取+汇总，不写文件，返回 (成功, 异常信息, total_dict)"""
    try:
        total, file_count = read_and_aggregate(folder, PROJECT_ROOT)
        return True, f"OK: {file_count} files, {len(total)} customers", total
    except Exception as e:
        return False, f"{type(e).__name__}: {e}\n{traceback.format_exc()}", {}


# ============================================================
#  一、Excel 异常输入测试
# ============================================================

def test_category_1():
    print("\n" + "=" * 60)
    print("  一、Excel 异常输入测试")
    print("=" * 60)

    # --- 1.1 正常 Excel ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "1_normal.xlsx"),
        ["日期", "客户", "金额", "备注"],
        [["2024-01-01", "公司A", 1000.00, "测试"],
         ["2024-01-02", "公司B", 2500.50, "测试"],
         ["2024-01-03", "公司A", 300.00, "测试"]],
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("1.1", "Excel输入", "正常Excel", ok, detail)

    # --- 1.2 空 Excel（0 行数据） ---
    reset_test_data()
    wb = Workbook()
    wb.save(os.path.join(TEST_DATA, "2_empty.xlsx"))
    wb.close()
    ok, detail = run_pipeline(TEST_DATA)
    # 预期：能跑完（file_count=1但无有效数据），不崩溃
    passed = ok  # 不崩溃即为通过
    record("1.2", "Excel输入", "空Excel（无任何内容）", passed, detail)

    # --- 1.3 只有表头无数据 ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "3_header_only.xlsx"),
        ["客户", "金额", "备注"],
        [],
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("1.3", "Excel输入", "只有表头无数据", ok, detail)

    # --- 1.4 缺少客户列 ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "4_no_customer.xlsx"),
        ["日期", "金额", "备注"],
        [["2024-01-01", 1000.00, "测试"]],
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("1.4", "Excel输入", "缺少客户列", ok, detail)

    # --- 1.5 缺少金额列 ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "5_no_money.xlsx"),
        ["日期", "客户", "备注"],
        [["2024-01-01", "公司A", "测试"]],
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("1.5", "Excel输入", "缺少金额列", ok, detail)

    # --- 1.6 金额包含文字 ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "6_money_text.xlsx"),
        ["客户", "金额"],
        [["公司A", 1000.00],
         ["公司B", "壹仟元整"],  # 中文金额
         ["公司C", "N/A"],        # 非数字
         ["公司A", 500.00],
         ["公司D", ""],           # 空金额
         ["公司E", None]],        # None 金额
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("1.6", "Excel输入", "金额包含文字/空值", ok, detail)

    # --- 1.7 多余列 ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "7_extra_cols.xlsx"),
        ["编号", "日期", "客户", "品名", "规格", "数量", "单价", "金额", "备注", "经手人"],
        [["001", "2024-01-01", "公司A", "产品X", "A1", 10, 100, 1000, "无", "张三"],
         ["002", "2024-01-02", "公司B", "产品Y", "B2", 5, 500, 2500, "无", "李四"]],
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("1.7", "Excel输入", "多余列（10列）", ok, detail)

    # --- 1.8 列顺序变化 ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "8_col_order.xlsx"),
        ["金额", "日期", "备注", "客户"],  # 金额在前，客户在后
        [[1000.00, "2024-01-01", "测试", "公司A"],
         [2500.00, "2024-01-02", "测试", "公司B"]],
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("1.8", "Excel输入", "列顺序变化", ok, detail)

    # --- 1.9 多个 Excel 格式不同 ---
    reset_test_data()
    # 格式A：标准格式
    make_workbook(
        os.path.join(TEST_DATA, "9a_format_a.xlsx"),
        ["客户", "金额"],
        [["公司A", 1000.00]],
    )
    # 格式B：客户叫"购货单位"，金额叫"价税合计"
    make_workbook(
        os.path.join(TEST_DATA, "9b_format_b.xlsx"),
        ["日期", "购货单位", "品名", "价税合计"],
        [["2024-01-01", "公司B", "产品X", 2500.00]],
    )
    # 格式C：客户叫"单位名称"，金额叫"合计金额"
    make_workbook(
        os.path.join(TEST_DATA, "9c_format_c.xlsx"),
        ["单位名称", "合计金额", "备注"],
        [["公司A", 800.00, "测试"]],
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("1.9", "Excel输入", "多个Excel格式不同", ok, detail)


# ============================================================
#  二、文件异常测试
# ============================================================

def test_category_2():
    print("\n" + "=" * 60)
    print("  二、文件异常测试")
    print("=" * 60)

    # --- 2.1 文件夹为空 ---
    reset_test_data()
    ok, detail = run_pipeline(TEST_DATA)
    record("2.1", "文件异常", "文件夹为空", ok, detail)

    # --- 2.2 文件夹不存在 ---
    non_exist = os.path.join(TEST_DATA, "_does_not_exist_")
    try:
        total, file_count = read_and_aggregate(non_exist, PROJECT_ROOT)
        record("2.2", "文件异常", "文件夹不存在",
               False,
               f"应抛出异常但没有！file_count={file_count}",
               "FileNotFoundError 未触发（os.listdir 应先报错）")
    except FileNotFoundError:
        record("2.2", "文件异常", "文件夹不存在", True, "正确抛出 FileNotFoundError")
    except Exception as e:
        record("2.2", "文件异常", "文件夹不存在", True,
               f"抛出 {type(e).__name__}: {e}")

    # --- 2.3 Excel 损坏 ---
    reset_test_data()
    # 创建一个伪装的 xlsx（实际是文本文件）
    fake_path = os.path.join(TEST_DATA, "3_corrupt.xlsx")
    with open(fake_path, "w", encoding="utf-8") as f:
        f.write("这不是一个真正的 Excel 文件")
    ok, detail = run_pipeline(TEST_DATA)
    record("2.3", "文件异常", "Excel损坏（非xlsx格式）", ok, detail)

    # --- 2.4 Excel 正在被打开（模拟 PermissionError） ---
    reset_test_data()
    # 创建一个正常 Excel，然后用只读打开锁定它
    normal_file = os.path.join(TEST_DATA, "4_locked.xlsx")
    make_workbook(normal_file, ["客户", "金额"], [["公司A", 1000.00]])

    # 模拟：先打开文件，再尝试读取（在 Windows 上 openpyxl 以只读方式打开，通常不会冲突）
    # 但写入时会遇到问题，我们测试写入冲突
    ok = True
    detail = ""

    total, fc = read_and_aggregate(TEST_DATA, PROJECT_ROOT)
    if fc > 0 and len(total) > 0:
        sorted_data = sort_by_amount(total)
        output_path = os.path.join(TEST_DATA, OUTPUT_FILE)

        # 先创建并占用输出文件
        lock_fd = open(output_path, "w")
        lock_fd.write("locked")
        lock_fd.flush()
        try:
            generate_summary(sorted_data, total, output_path)
            ok = False
            detail = "应抛出 PermissionError 但没有"
        except PermissionError:
            ok = True
            detail = "正确抛出 PermissionError"
        except Exception as e:
            ok = True
            detail = f"抛出 {type(e).__name__}: {e}（非 PermissionError 但也没崩溃）"
        finally:
            lock_fd.close()

    record("2.4", "文件异常", "输出文件被占用", ok, detail)

    # --- 2.5 输出文件已存在（覆盖测试） ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "5_existing.xlsx"),
        ["客户", "金额"],
        [["公司A", 1000.00]],
    )
    output_path = os.path.join(TEST_DATA, "5_output_test.xlsx")
    # 先写一个旧输出
    wb = Workbook()
    wb.active.append(["旧数据", "旧数据"])
    wb.save(output_path)
    wb.close()

    total, fc = read_and_aggregate(TEST_DATA, PROJECT_ROOT)
    if fc > 0 and len(total) > 0:
        sorted_data = sort_by_amount(total)
        try:
            generate_summary(sorted_data, total, output_path)
            # 验证覆盖成功
            verify_wb = load_workbook(output_path)
            first_cell = verify_wb.active["A1"].value
            verify_wb.close()
            if first_cell == "客户":
                ok, detail = True, "覆盖成功，旧文件被替换"
            else:
                ok, detail = False, f"覆盖后A1应为'客户'，实际为'{first_cell}'"
        except Exception as e:
            ok, detail = False, f"{type(e).__name__}: {e}"
    else:
        ok, detail = False, "无法读取测试文件"

    record("2.5", "文件异常", "输出文件已存在（覆盖）", ok, detail)


# ============================================================
#  三、大数据测试
# ============================================================

def test_category_3():
    print("\n" + "=" * 60)
    print("  三、大数据测试")
    print("=" * 60)

    reset_test_data()
    NUM_FILES = 10
    ROWS_PER_FILE = 10000
    CUSTOMERS = ["客户A", "客户B", "客户C", "客户D", "客户E",
                 "客户F", "客户G", "客户H", "客户I", "客户J"]

    print(f"  生成 {NUM_FILES} 个文件，每个 {ROWS_PER_FILE:,} 行...")

    gen_start = time.perf_counter()
    for i in range(NUM_FILES):
        filepath = os.path.join(TEST_DATA, f"big_{i+1:02d}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.append(["客户", "金额", "备注"])
        for j in range(ROWS_PER_FILE):
            customer = CUSTOMERS[j % len(CUSTOMERS)]
            amount = round(100 + (j * 137.5) % 90000, 2)
            # 每500行加入一些特殊值
            if j % 500 == 0:
                ws.append([customer, amount, f"记录{j}"])
            else:
                ws.append([customer, amount])
        wb.save(filepath)
        wb.close()
        if (i + 1) % 3 == 0:
            print(f"    已生成 {i+1}/{NUM_FILES} 个文件...")

    gen_elapsed = time.perf_counter() - gen_start
    print(f"  生成耗时: {gen_elapsed:.1f}s")

    # --- 读取测试 ---
    print(f"\n  开始读取并汇总...")
    mem_before = _get_memory_mb()
    read_start = time.perf_counter()

    total, file_count = read_and_aggregate(TEST_DATA, PROJECT_ROOT)

    read_elapsed = time.perf_counter() - read_start
    mem_after = _get_memory_mb()

    read_ok = file_count == NUM_FILES and len(total) > 0

    detail_str = (
        f"文件数: {file_count}/{NUM_FILES}, "
        f"客户数: {len(total)}, "
        f"耗时: {read_elapsed:.1f}s, "
        f"内存 Δ: {mem_after - mem_before:+.1f}MB"
    )

    record("3.1", "大数据", f"读取 {NUM_FILES}×{ROWS_PER_FILE:,}行", read_ok, detail_str)

    if not read_ok:
        return  # 后续测试无意义

    # --- 排序测试 ---
    sort_start = time.perf_counter()
    sorted_data = sort_by_amount(total)
    sort_elapsed = time.perf_counter() - sort_start
    record("3.2", "大数据", "排序", sort_elapsed < 5.0,
           f"耗时: {sort_elapsed:.3f}s, 条目: {len(sorted_data)}")

    # --- 写入测试 ---
    output_path = os.path.join(TEST_DATA, "big_output.xlsx")
    write_start = time.perf_counter()
    try:
        generate_summary(sorted_data, total, output_path)
        write_elapsed = time.perf_counter() - write_start
        file_size_mb = os.path.getsize(output_path) / (1024 * 1024)
        record("3.3", "大数据", "写入汇总", True,
               f"耗时: {write_elapsed:.1f}s, 文件大小: {file_size_mb:.2f}MB")
    except Exception as e:
        record("3.3", "大数据", "写入汇总", False,
               f"{type(e).__name__}: {e}\n{traceback.format_exc()}")

    # --- 内存检查 ---
    mem_ok = (mem_after - mem_before) < 500  # 不允许增长超过 500MB
    record("3.4", "大数据", "内存检查", mem_ok,
           f"内存增长 {mem_after - mem_before:+.1f}MB (阈值 500MB)")

    total_elapsed = time.perf_counter() - read_start
    record("3.5", "大数据", "总耗时", total_elapsed < 60.0,
           f"总耗时: {total_elapsed:.1f}s (阈值 60s)")

    # 打印总体统计
    print(f"\n  大数据测试总结:")
    print(f"    总耗时: {total_elapsed:.1f}s")
    print(f"    内存增长: {mem_after - mem_before:+.1f}MB")
    print(f"    处理行数: {NUM_FILES * ROWS_PER_FILE:,}")
    print(f"    客户数: {len(total)}")


def _get_memory_mb():
    """获取当前进程内存占用（MB），跨平台"""
    try:
        import psutil
        return psutil.Process(os.getpid()).memory_info().rss / (1024 * 1024)
    except ImportError:
        return -1  # psutil 未安装


# ============================================================
#  四、额外边界测试
# ============================================================

def test_category_4():
    """补充边界条件测试"""
    print("\n" + "=" * 60)
    print("  四、补充边界条件测试")
    print("=" * 60)

    # --- 4.1 客户名为空但金额有效 ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "b1_empty_customer.xlsx"),
        ["客户", "金额"],
        [[None, 1000.00],     # None 客户
         ["", 2000.00],       # 空字符串客户
         ["正常客户", 500.00]],  # 正常
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("4.1", "边界条件", "客户名为空/None", ok, detail)

    # --- 4.2 客户名重复（去重汇总） ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "b2_duplicate.xlsx"),
        ["客户", "金额"],
        [["公司A", 100.00],
         ["公司A", 200.00],
         ["公司A", 300.00],
         ["公司B", 500.00]],
    )
    ok, detail, total = capture_read_only(TEST_DATA)
    if ok:
        sum_a = total.get("公司A", 0)
        sum_b = total.get("公司B", 0)
        if sum_a == 600.00 and sum_b == 500.00:
            detail += f" | 汇总正确: A=600, B=500"
        else:
            ok = False
            detail += f" | 汇总错误: A={sum_a}, B={sum_b}"
    record("4.2", "边界条件", "客户名重复汇总", ok, detail)

    # --- 4.3 含有 merged cells 的 Excel ---
    reset_test_data()
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额", "备注"])
    ws.append(["公司A", 1000.00, "正常行"])
    ws.merge_cells("A3:B3")
    ws["A3"] = "合并单元格测试"
    ws.append(["公司B", 2000.00, "另一行"])
    wb.save(os.path.join(TEST_DATA, "b3_merged.xlsx"))
    wb.close()
    ok, detail = run_pipeline(TEST_DATA)
    record("4.3", "边界条件", "合并单元格", ok, detail)

    # --- 4.4 含有公式的金额列 ---
    reset_test_data()
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["公司A", 1000.00])
    ws["B3"] = "=B2*2"  # 公式
    ws["A3"] = "公司B"
    wb.save(os.path.join(TEST_DATA, "b4_formula.xlsx"))
    wb.close()
    ok, detail = run_pipeline(TEST_DATA)
    record("4.4", "边界条件", "金额为公式", ok, detail)

    # --- 4.5 表头有空格 ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "b5_whitespace_header.xlsx"),
        ["  客户  ", "  金额  "],  # 带空格
        [["公司A", 1000.00]],
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("4.5", "边界条件", "表头含首尾空格", ok, detail)

    # --- 4.6 数值类型异常（bool, date) ---
    reset_test_data()
    wb = Workbook()
    ws = wb.active
    ws.append(["客户", "金额"])
    ws.append(["公司A", 1000.00])
    ws.append(["公司B", True])    # bool 值
    ws.append(["公司C", 2000.00])
    wb.save(os.path.join(TEST_DATA, "b6_bool_amount.xlsx"))
    wb.close()
    ok, detail = run_pipeline(TEST_DATA)
    record("4.6", "边界条件", "金额为布尔值", ok, detail, None)

    # --- 4.7 大数字/负数 ---
    reset_test_data()
    make_workbook(
        os.path.join(TEST_DATA, "b7_large_numbers.xlsx"),
        ["客户", "金额"],
        [["公司A", 999999999999.99],
         ["公司B", -500.00],
         ["公司C", 0.00],
         ["公司D", 0.01]],
    )
    ok, detail = run_pipeline(TEST_DATA)
    record("4.7", "边界条件", "大数字/负数/零值", ok, detail)

    # --- 4.8 仅有临时文件（~$ 开头） ---
    reset_test_data()
    # 创建一个正常文件和一个 ~$ 文件
    make_workbook(
        os.path.join(TEST_DATA, "~$temp_file.xlsx"),
        ["客户", "金额"],
        [["公司A", 1000.00]],
    )
    ok, detail = run_pipeline(TEST_DATA)
    # 应该报告没有文件（因为 ~$ 被过滤掉）
    record("4.8", "边界条件", "仅含~$临时文件", ok,
           f"应报告无有效文件: {detail}")


# ============================================================
#  主入口
# ============================================================

def print_header():
    print("=" * 60)
    print("  Excel 自动汇总助手 v1.2.2.1 — 稳定性测试")
    print(f"  测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Python:    {sys.version}")
    print("=" * 60)


def print_summary():
    print("\n\n" + "=" * 60)
    print("  测试结果汇总")
    print("=" * 60)

    passed = [r for r in RESULTS if r[3]]
    failed = [r for r in RESULTS if not r[3]]

    total = len(RESULTS)
    pass_count = len(passed)
    fail_count = len(failed)

    print(f"\n  总计: {total} 项")
    print(f"  通过: {pass_count} 项")
    print(f"  失败: {fail_count} 项")

    if fail_count > 0:
        print(f"\n  ❌ 失败项明细:")
        for r in failed:
            print(f"     [{r[0]}] {r[2]}: {r[4][:100]}")

    # 按分类汇总
    print(f"\n  按分类:")
    categories = {}
    for r in RESULTS:
        cat = r[1]
        if cat not in categories:
            categories[cat] = {"pass": 0, "fail": 0}
        if r[3]:
            categories[cat]["pass"] += 1
        else:
            categories[cat]["fail"] += 1

    for cat, counts in categories.items():
        total_cat = counts["pass"] + counts["fail"]
        bar = "█" * counts["pass"] + "░" * counts["fail"]
        print(f"    {cat}: {bar} {counts['pass']}/{total_cat}")

    return pass_count, fail_count, total


def save_report():
    """输出详细报告到文件"""
    report_path = os.path.join(PROJECT_ROOT, "tests", "test_report.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write("  Excel 自动汇总助手 v1.2.2.1 — 稳定性测试报告\n")
        f.write(f"  测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")

        for r in RESULTS:
            test_id, category, sub_category, passed, detail, crash_tb = r
            status = "PASS" if passed else "FAIL"
            f.write(f"[{status}] {test_id} : {sub_category}\n")
            if detail:
                f.write(f"  {detail}\n")
            if crash_tb:
                f.write(f"  TRACEBACK:\n")
                for line in crash_tb.strip().split("\n"):
                    f.write(f"    {line}\n")
            f.write("\n")

        passed = sum(1 for r in RESULTS if r[3])
        failed = sum(1 for r in RESULTS if not r[3])
        f.write("-" * 60 + "\n")
        f.write(f"总计: {len(RESULTS)} 项 | 通过: {passed} | 失败: {failed}\n")

    print(f"\n  详细报告已保存到: {report_path}")


if __name__ == "__main__":
    print_header()

    test_category_1()
    test_category_2()
    test_category_3()
    test_category_4()

    pass_count, fail_count, total = print_summary()
    save_report()

    # 返回退出码：有失败则返回 1
    sys.exit(0 if fail_count == 0 else 1)
