import os
from openpyxl import load_workbook
import xlrd
from .config import ERROR_LOG, CUSTOMER_KEYWORDS, MONEY_KEYWORDS, CUSTOM_HEADER_MAP
from .field_mapper import (
    detect_columns, get_header_list, diagnose_failure, ColumnMapping,
    FieldDetection, _get_effective_config, _match_all_fields,
    missing_required_fields,
)

# 支持的文件扩展名
_SUPPORTED_EXTENSIONS = (".xlsx", ".xlsm", ".xls")


def _print_field_error(filename, headers, cust_col, money_col, is_empty):
    """输出结构化的字段识别错误信息。

    Args:
        filename: 文件名（仅用于显示）
        headers: 实际表头列表
        cust_col: 客户列索引，未找到为 None
        money_col: 金额列索引，未找到为 None
        is_empty: 表头是否为空
    """
    # ---- 情况 1：表头为空 ----
    if is_empty:
        print(f"  警告：文件 {filename} 缺少列名（第 1 行为空），已跳过")
        print(f"    请在 Excel 第 1 行添加列名，例如：日期、客户、金额")
        print("-" * 35)
        return

    # ---- 情况 2：部分或全部缺失 ----
    print(f"  警告：文件 {filename} 列名识别失败，已跳过")
    print(f"    检测到的列名：{headers}")

    # 逐字段报告
    _print_field_status("客户列", cust_col, headers, CUSTOMER_KEYWORDS)
    _print_field_status("金额列", money_col, headers, MONEY_KEYWORDS)

    # ---- 解决方案 ----
    print("  " + "-" * 30)
    print("  解决方法（选一种即可）：")
    print(f"    ① 修改 Excel 表头")
    if cust_col is None:
        print(f"       将客户列的表头改为：{' / '.join(CUSTOMER_KEYWORDS[:4])}")
    if money_col is None:
        print(f"       将金额列的表头改为：{' / '.join(MONEY_KEYWORDS[:4])}")

    # 动态生成自定义映射建议
    clean = [str(h).strip() if h is not None else "" for h in headers]
    map_entries = []
    if cust_col is None:
        for h in clean:
            if h and h not in [str(x) for x in CUSTOM_HEADER_MAP]:
                map_entries.append(f'"{h}": "customer"')
                break
    if money_col is None:
        for h in clean:
            if h and h not in [str(x) for x in CUSTOM_HEADER_MAP]:
                # 避免和客户列选同一个
                already = map_entries[0].split(":")[0].strip('"') if map_entries else ""
                if h != already:
                    map_entries.append(f'"{h}": "money"')
                    break

    if map_entries:
        print(f"    ② 在程序设置文件（config.py）的 CUSTOM_HEADER_MAP 中添加：")
        print(f"         {{{', '.join(map_entries)}}}")
    else:
        print(f"    ② 在程序设置文件（config.py）的 CUSTOM_HEADER_MAP 中添加自定义映射")

    print(f"    ③ 在程序设置文件（config.py）中扩充关键词列表")
    print("  " + "-" * 30)


def _print_field_status(label, col_index, headers, keywords):
    """输出单个字段的识别状态。"""
    if col_index is not None:
        header_name = headers[col_index] if col_index < len(headers) else "?"
        print(f"    [OK] {label} -> 第 {col_index + 1} 列（\"{header_name}\"）")
    else:
        print(f"    [MISS] {label} -> 未找到匹配的列名")
        print(f"       支持的表头名称：{' / '.join(keywords)}")


# ============================================================
#  v1.4: 多格式支持 — 内部辅助
# ============================================================

def _get_excel_files(folder):
    """扫描目录，返回有效 Excel 文件列表（过滤临时/隐藏/锁定文件）。

    Returns:
        list[str]: 文件名列表（仅名称，不含路径）
    """
    try:
        all_items = os.listdir(folder)
    except PermissionError:
        return []

    result = []
    for f in all_items:
        # 过滤临时文件（~$ 开头）
        if f.startswith("~$"):
            continue
        # 过滤隐藏文件
        if f.startswith("."):
            continue
        # 过滤非 Excel 扩展名
        lower = f.lower()
        if not lower.endswith(_SUPPORTED_EXTENSIONS):
            continue
        # 过滤空文件（文件锁常表现为 0 字节）
        fpath = os.path.join(folder, f)
        try:
            if os.path.getsize(fpath) == 0:
                continue
        except OSError:
            continue
        result.append(f)
    return result


def _detect_columns_xls(sheet, header_map=None):
    """xlrd 版字段检测 — 使用 field_mapper 的动态字段匹配。

    Args:
        sheet: xlrd 工作表对象
        header_map: 当前任务人工映射 {表头: 字段名}，可为 None

    Returns:
        list[FieldDetection] 或 None
    """
    if sheet.nrows < 1:
        return None
    headers = sheet.row_values(0)
    if all(h == "" or h is None for h in headers):
        return None

    clean = [str(h).strip() if h is not None and str(h).strip() else "" for h in headers]
    config = _get_effective_config()
    detections = _match_all_fields(clean, config, header_map)

    # 必须识别到前两个字段（分组键 + 求和字段），避免位置错位
    required = [f["name"] for f in config[:2]]
    detected = {d.field_name for d in detections}
    if not all(name in detected for name in required):
        return None
    return detections


def _match_in_headers(headers, keywords):
    """两级匹配：精确匹配 → 包含匹配。"""
    for i, header in enumerate(headers):
        if not header:
            continue
        for kw in keywords:
            if header == kw:
                return i
    for i, header in enumerate(headers):
        if not header:
            continue
        for kw in keywords:
            if kw in header:
                return i
    return None


def _match_custom_in_headers(headers, field_type):
    """在自定义映射中查找匹配的列。"""
    for i, header in enumerate(headers):
        if not header:
            continue
        if header in CUSTOM_HEADER_MAP and CUSTOM_HEADER_MAP[header] == field_type:
            return i
    return None


def _add_row_to_total(total, row_values, det_map, group_field, sum_field, filename):
    """验证单行数据并累加到汇总字典（v1.5.2 Hotfix1：按字段名 + 列索引读取）。

    Args:
        total: {group_key: {field_name: value, ...}}
        row_values: 行数据列表
        det_map: {field_name: col_index} 字段名 → 列索引映射
        group_field: 分组键字段名（配置第 1 个字段）
        sum_field: 求和字段名（配置第 2 个字段）
        filename: 文件名（用于警告）
    """
    if not group_field or not sum_field:
        return

    # 提取分组键（按字段名 → 列索引读取，避免位置错位）
    group_col = det_map.get(group_field)
    if group_col is None or group_col >= len(row_values):
        return
    key_raw = row_values[group_col]

    # 验证分组键
    if key_raw is None:
        return
    if isinstance(key_raw, str):
        key_raw = key_raw.strip()
        if not key_raw:
            return
    key = str(key_raw)

    # 提取求和字段（按字段名 → 列索引读取）
    sum_col = det_map.get(sum_field)
    if sum_col is None or sum_col >= len(row_values):
        return
    money_raw = row_values[sum_col]

    # 金额：字符串 → 数字
    if isinstance(money_raw, str):
        money_stripped = money_raw.strip()
        if not money_stripped:
            return
        try:
            money_raw = float(money_stripped)
        except ValueError:
            print(f"  警告：文件 {filename} 中客户 [{key}] 的金额 [{money_raw}] 无法识别为数字，已跳过")
            return

    if isinstance(money_raw, bool):
        print(f"  警告：文件 {filename} 中客户 [{key}] 的金额格式无法识别，已跳过")
        return

    if not isinstance(money_raw, (int, float)):
        print(f"  警告：文件 {filename} 中客户 [{key}] 的金额 [{money_raw}] 不是数字，已跳过")
        return

    # 构建行数据（按字段名 → 列索引读取，而非按字段列表位置）
    row_data = {}
    for field_name, col in det_map.items():
        if col < len(row_values):
            row_data[field_name] = row_values[col]
        else:
            row_data[field_name] = ""

    if key not in total:
        total[key] = dict(row_data)
        # 确保求和字段初始为数值
        total[key][sum_field] = money_raw
    else:
        total[key][sum_field] += money_raw


def read_and_aggregate(folder, base_dir, file_names=None, field_mappings=None):
    """读取文件夹内 Excel 文件（.xlsx/.xlsm/.xls），按配置字段汇总。

    向后兼容封装：返回 (total, file_count, field_names)。实际逻辑见
    aggregate_with_report（含跳过文件报告，用于核心字段完整性保护）。

    v1.5.0: 根据 field_config.json 动态检测字段，不再硬编码客户/金额。
    v1.5.2 Hotfix3: 新增 file_names 参数，只汇总指定文件名；为 None 时汇总全部。
    v1.5.4: 新增 field_mappings（当前任务人工映射）。

    Args:
        folder: 文件夹路径
        base_dir: 日志目录
        file_names: 需要汇总的文件名列表（仅名称）；None 表示汇总全部
        field_mappings: {文件名: {表头: 字段名}}，当前任务人工映射

    Returns:
        (total, file_count, field_names):
            - total: {分组键: {字段名: 值, ...}}
            - file_count: 有效文件数
            - field_names: 字段名列表（与配置顺序一致）
    """
    report = aggregate_with_report(folder, base_dir, file_names, field_mappings)
    return report["total"], report["file_count"], report["field_names"]


def aggregate_with_report(folder, base_dir, file_names=None, field_mappings=None):
    """读取文件夹内 Excel 文件（.xlsx/.xlsm/.xls），按配置字段汇总，并报告跳过文件。

    v1.5.4 核心字段完整性保护：当某个选中文件的核心字段（分组键 + 求和字段）
    无法识别时，不再静默跳过该文件，而是将其记录到 skipped 报告中，由调用方
    （GUI 的 WorkerThread）决定是否阻止生成汇总文件。

    Args:
        folder: 文件夹路径
        base_dir: 日志目录
        file_names: 需要汇总的文件名列表（仅名称）；None 表示汇总全部
        field_mappings: {文件名: {表头: 字段名}}，当前任务人工映射（最高优先级）

    Returns:
        dict:
            total: {分组键: {字段名: 值, ...}}
            file_count: 参与汇总的有效文件数
            field_names: 字段名列表（与配置顺序一致）
            processed_count: 成功处理的文件数
            skipped: [{file, reason, missing}]，reason ∈ {"core_field", "unreadable"}
    """
    excel_files = _get_excel_files(folder)
    if file_names is not None:
        allowed = set(file_names)
        excel_files = [f for f in excel_files if f in allowed]
    total = {}
    config = _get_effective_config()
    field_names = [f["name"] for f in config]
    # 分组键 = 配置第 0 个字段，求和字段 = 配置第 1 个字段（v1.5.2 Hotfix1）
    group_field = field_names[0] if field_names else None
    sum_field = field_names[1] if len(field_names) >= 2 else None

    skipped = []
    processed = 0

    for file in excel_files:
        file_path = os.path.join(folder, file)
        ext = os.path.splitext(file)[1].lower()
        header_map = (field_mappings or {}).get(file) or None

        # .xlsx / .xlsm — openpyxl
        if ext in (".xlsx", ".xlsm"):
            try:
                wb = load_workbook(file_path, data_only=True)
            except Exception as e:
                print(f"  警告：文件 {file} 无法读取，已跳过")
                with open(os.path.join(base_dir, ERROR_LOG), "a", encoding="utf-8") as logf:
                    logf.write(f"无法打开 {file_path}：{repr(e)}\n")
                skipped.append({"file": file, "reason": "unreadable", "missing": []})
                continue

            try:
                ws = wb.active
                detections = detect_columns(ws, header_map)
                if detections is None:
                    headers, cust_col, money_col, is_empty = diagnose_failure(ws, header_map)
                    _print_field_error(file, headers, cust_col, money_col, is_empty)
                    with open(os.path.join(base_dir, ERROR_LOG), "a", encoding="utf-8") as logf:
                        logf.write(f"字段识别失败 {file_path}：表头={headers}\n")
                    clean = [str(h).strip() if h is not None else "" for h in headers]
                    missing = missing_required_fields(clean, header_map)
                    skipped.append({"file": file, "reason": "core_field", "missing": missing})
                    continue

                det_map = {d.field_name: d.col_index for d in detections}
                for row in ws.iter_rows(min_row=2, values_only=True):
                    _add_row_to_total(total, row, det_map, group_field, sum_field, file)
                processed += 1
            finally:
                wb.close()

        # .xls — xlrd
        elif ext == ".xls":
            try:
                wb_xls = xlrd.open_workbook(file_path)
            except Exception as e:
                print(f"  警告：文件 {file} 无法读取，已跳过")
                with open(os.path.join(base_dir, ERROR_LOG), "a", encoding="utf-8") as logf:
                    logf.write(f"无法打开 {file_path}：{repr(e)}\n")
                skipped.append({"file": file, "reason": "unreadable", "missing": []})
                continue

            try:
                sheet = wb_xls.sheet_by_index(0)
                detections = _detect_columns_xls(sheet, header_map)
                if detections is None:
                    headers = sheet.row_values(0)
                    print(f"  警告：文件 {file} 列名识别失败，已跳过")
                    print(f"    检测到的列名：{headers}")
                    with open(os.path.join(base_dir, ERROR_LOG), "a", encoding="utf-8") as logf:
                        logf.write(f"字段识别失败 {file_path}：表头={headers}\n")
                    clean = [str(h).strip() if h is not None else "" for h in headers]
                    missing = missing_required_fields(clean, header_map)
                    skipped.append({"file": file, "reason": "core_field", "missing": missing})
                    continue

                det_map = {d.field_name: d.col_index for d in detections}
                for row_idx in range(1, sheet.nrows):
                    row = sheet.row_values(row_idx)
                    _add_row_to_total(total, row, det_map, group_field, sum_field, file)
                processed += 1
            finally:
                wb_xls.release_resources()

    return {
        "total": total,
        "file_count": len(excel_files),
        "field_names": field_names,
        "processed_count": processed,
        "skipped": skipped,
    }
