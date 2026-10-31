import os
import json
from dataclasses import dataclass, field
from .config import CUSTOMER_KEYWORDS, MONEY_KEYWORDS, CUSTOM_HEADER_MAP

SETTINGS_FILENAME = "settings.json"
FIELD_CONFIG_FILENAME = "field_config.json"

# ============================================================
#  v1.5.0: 默认字段配置（list-of-objects 格式）
#  字段顺序决定输出列顺序；第 0 个字段为分组键，第 1 个为求和字段
# ============================================================
_DEFAULT_FIELD_CONFIG = [
    {
        "name": "客户",
        "keywords": [
            "客户", "客户名称", "购货单位", "单位名称",
            "公司", "顾客", "收货人", "买方",
        ],
    },
    {
        "name": "金额",
        "keywords": [
            "金额", "总金额", "合计金额", "价税合计",
            "货款", "费用", "小计", "应收金额",
        ],
    },
]

# 内存缓存：当前生效的字段配置
_FIELD_CONFIG = []


@dataclass
class ColumnMapping:
    """列索引映射：记录客户列和金额列在工作表中的位置（0-based）。

    v1.5.0 保留此类以兼容旧接口；新代码使用 FieldDetection。
    """
    customer_col: int
    money_col: int


@dataclass
class FieldDetection:
    """单个字段的检测结果：字段名 → 列索引。"""
    field_name: str
    col_index: int


def _get_default_field_names():
    """返回默认字段名列表（前两个用于关键词替换兼容）。"""
    return [f["name"] for f in _DEFAULT_FIELD_CONFIG]


def _field_config_to_dict(config_list):
    """将 list-of-objects 格式转为旧 dict 格式（向后兼容）。"""
    result = {}
    for f in config_list:
        name = f.get("name", "")
        keywords = f.get("keywords", [])
        if name and keywords:
            result[name] = list(keywords)
    return result


def detect_columns(ws, header_map=None):
    """从工作表表头行自动识别所有已配置字段。

    匹配策略（优先级从高到低，v1.5.4 加入当前任务人工映射）：
        0. 当前任务人工映射（header_map，最高优先级）
        1. 用户自定义映射（CUSTOM_HEADER_MAP）
        2. 字段名称精确匹配
        3. 字段名称包含匹配
        4. 关键词精确匹配
        5. 关键词包含匹配

    Args:
        ws: openpyxl 工作表对象
        header_map: 当前任务人工映射 {表头: 字段名}，可为 None

    Returns:
        list[FieldDetection] 若至少识别到第 0、1 个字段，
        None 若前两个必需字段无法全部识别
    """
    headers = get_header_list(ws)

    if all(h is None for h in headers):
        return None

    clean = [str(h).strip() if h is not None else "" for h in headers]
    config = _get_effective_config()

    detections = _match_all_fields(clean, config, header_map)

    # 必须识别到前两个字段（分组键 + 求和字段），否则视为识别失败。
    # 只检查 len(detections) >= 2 会在中间字段缺失时造成位置错位：
    # 例如配置 [客户, 金额, 电话] 中「客户」缺失时，detections 会变成
    # [金额, 电话]，导致金额被当作分组键、电话被当作求和字段。
    required = [f["name"] for f in config[:2]]
    detected = {d.field_name for d in detections}
    if not all(name in detected for name in required):
        return None

    return detections


def _get_effective_config():
    """返回当前生效的字段配置列表。"""
    if _FIELD_CONFIG:
        return _FIELD_CONFIG
    return _DEFAULT_FIELD_CONFIG


def _match_all_fields(headers, field_config, header_map=None):
    """对所有配置字段进行匹配，返回 FieldDetection 列表。

    每个表头列最多匹配一个字段（先匹配先得）。
    匹配顺序：当前任务人工映射 → 字段名精确 → 字段名包含 → 关键词精确 → 关键词包含。
    """
    used_cols = set()
    detections = []

    for field_def in field_config:
        field_name = field_def.get("name", "")
        keywords = field_def.get("keywords", [])

        match_idx = _match_field_in_headers(headers, field_name, keywords, used_cols, header_map)
        if match_idx is not None:
            detections.append(FieldDetection(field_name=field_name, col_index=match_idx))
            used_cols.add(match_idx)

    return detections


def _match_field_in_headers(headers, field_name, keywords, used_cols, header_map=None):
    """在表头中匹配一个字段，返回列索引或 None。

    匹配优先级（v1.5.4）：
        0. 当前任务人工映射（header_map，最高优先级）
        1. 用户自定义映射（CUSTOM_HEADER_MAP）
        2. 字段名精确匹配
        3. 字段名包含匹配
        4. 关键词精确匹配
        5. 关键词包含匹配
    """
    # 0. 当前任务人工映射（最高优先级）
    if header_map:
        for i, h in enumerate(headers):
            if i in used_cols or not h:
                continue
            if h in header_map and header_map[h] == field_name:
                return i

    # 1. 用户自定义映射优先
    for i, h in enumerate(headers):
        if i in used_cols or not h:
            continue
        if h in CUSTOM_HEADER_MAP and CUSTOM_HEADER_MAP[h] == field_name:
            return i

    # 1. 字段名精确匹配
    for i, h in enumerate(headers):
        if i in used_cols or not h:
            continue
        if h == field_name:
            return i

    # 2. 字段名包含匹配
    for i, h in enumerate(headers):
        if i in used_cols or not h:
            continue
        if field_name in h:
            return i

    # 3. 关键词精确匹配
    for i, h in enumerate(headers):
        if i in used_cols or not h:
            continue
        for kw in keywords:
            if h == kw:
                return i

    # 4. 关键词包含匹配
    for i, h in enumerate(headers):
        if i in used_cols or not h:
            continue
        for kw in keywords:
            if kw in h:
                return i

    return None


def get_header_list(ws):
    """读取工作表第一行的表头列表"""
    return [cell.value for cell in ws[1]]


def diagnose_failure(ws, header_map=None):
    """详细诊断识别失败原因。

    Args:
        ws: openpyxl 工作表对象
        header_map: 当前任务人工映射 {表头: 字段名}，可为 None

    Returns:
        (headers, customer_col, money_col, is_empty):
            - headers: 原始表头列表
            - customer_col: 第 0 个字段（分组键）列索引，未找到为 None
            - money_col: 第 1 个字段（求和字段）列索引，未找到为 None
            - is_empty: 表头是否为空
    """
    headers = get_header_list(ws)
    if all(h is None for h in headers):
        return headers, None, None, True

    clean = [str(h).strip() if h is not None else "" for h in headers]
    config = _get_effective_config()
    detections = _match_all_fields(clean, config, header_map)

    customer_col = None
    money_col = None
    for d in detections:
        if customer_col is None:
            customer_col = d.col_index
        elif money_col is None:
            money_col = d.col_index
        else:
            break

    return headers, customer_col, money_col, False


def missing_required_fields(headers, header_map=None):
    """给定清洗后的表头列表，返回缺失的核心字段（前两个）名列表。

    用于 v1.5.4 核心字段完整性保护：当某文件的核心字段（分组键 + 求和字段）
    无法识别时，汇总应被阻止，而不是静默跳过该文件。

    Args:
        headers: 清洗后的表头字符串列表
        header_map: 当前任务人工映射 {表头: 字段名}，可为 None

    Returns:
        list[str]: 缺失的核心字段名（空列表表示核心字段全部可识别）
    """
    config = _get_effective_config()
    detections = _match_all_fields(headers, config, header_map)
    detected = {d.field_name for d in detections}
    required = [f["name"] for f in config[:2]]
    return [name for name in required if name not in detected]


def _match_column(headers, keywords):
    """两级匹配：精确匹配 → 包含匹配。"""
    # 第一级：精确匹配（header == keyword）
    for i, header in enumerate(headers):
        if not header:
            continue
        for kw in keywords:
            if header == kw:
                return i

    # 第二级：包含匹配（keyword in header）
    for i, header in enumerate(headers):
        if not header:
            continue
        for kw in keywords:
            if kw in header:
                return i

    return None


def _match_custom(headers, field_type):
    """在自定义映射中查找匹配的列。

    Args:
        headers: 清洗后的表头字符串列表
        field_type: "customer" 或 "money"

    Returns:
        匹配到的列索引（0-based），未匹配返回 None
    """
    for i, header in enumerate(headers):
        if not header:
            continue
        if header in CUSTOM_HEADER_MAP and CUSTOM_HEADER_MAP[header] == field_type:
            return i
    return None


# ============================================================
#  v1.3 新增：内部辅助 — JSON 文件读写
# ============================================================

def _settings_path(base_dir):
    """返回 settings.json 的完整路径。"""
    return os.path.join(base_dir, SETTINGS_FILENAME)


def _read_settings(base_dir):
    """读取 settings.json 返回 dict。不存在或损坏返回 {}。"""
    path = _settings_path(base_dir)
    if not os.path.isfile(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _write_settings(base_dir, settings):
    """写入 settings.json。"""
    path = _settings_path(base_dir)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


# ============================================================
#  v1.3 新增：字段映射持久化（公共接口）
# ============================================================

def load_settings(base_dir):
    """从 settings.json 读取 field_mapping 并合并到 CUSTOM_HEADER_MAP。

    启动时调用一次即可。settings.json 不存在时静默跳过。
    config.py 中已硬编码的映射不会被覆盖。

    Args:
        base_dir: 程序根目录（get_base_dir() 的返回值）
    """
    settings_path = os.path.join(base_dir, SETTINGS_FILENAME)
    if not os.path.isfile(settings_path):
        return

    try:
        with open(settings_path, "r", encoding="utf-8") as f:
            settings = json.load(f)
    except (json.JSONDecodeError, OSError):
        return

    field_mapping = settings.get("field_mapping", {})
    if not isinstance(field_mapping, dict):
        return

    # 合并到内存（后加载的 settings.json 覆盖 config.py 硬编码，
    # 因为用户通过 GUI 保存的映射是更"新"的配置）
    for header, field_type in field_mapping.items():
        CUSTOM_HEADER_MAP[header] = field_type


# ============================================================
#  v1.4.1 新增：字段规则配置（field_config.json）
# ============================================================

def load_field_config(base_dir):
    """从 field_config.json 读取用户自定义字段规则，并更新关键词列表。

    格式兼容：
        - 新格式 [{"name":"客户","keywords":[...]}, ...] → 直接使用
        - 旧格式 {"客户":[...], "金额":[...]} → 自动转换
        - 文件不存在 → 使用默认配置

    同时更新 CUSTOMER_KEYWORDS / MONEY_KEYWORDS（前两个字段的关键词）。

    Args:
        base_dir: 程序根目录
    """
    global _FIELD_CONFIG
    path = os.path.join(base_dir, FIELD_CONFIG_FILENAME)

    if not os.path.isfile(path):
        _FIELD_CONFIG = []
        return

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        _FIELD_CONFIG = []
        return

    # 格式检测与转换
    if isinstance(data, list):
        _FIELD_CONFIG = data
    elif isinstance(data, dict):
        # 旧格式 → 新格式
        _FIELD_CONFIG = [
            {"name": name, "keywords": kws}
            for name, kws in data.items()
        ]
    else:
        _FIELD_CONFIG = []

    # 更新关键词列表（前两个字段）
    _sync_keywords_from_config()


def _sync_keywords_from_config():
    """将当前字段配置的前两个字段的关键词同步到 CUSTOMER_KEYWORDS / MONEY_KEYWORDS。"""
    config = _get_effective_config()
    if len(config) >= 1:
        kws = config[0].get("keywords", [])
        if kws:
            _replace_list(CUSTOMER_KEYWORDS, kws)
    if len(config) >= 2:
        kws = config[1].get("keywords", [])
        if kws:
            _replace_list(MONEY_KEYWORDS, kws)


def _replace_list(target_list, new_items):
    """原地替换列表内容（保持引用不变）。"""
    target_list.clear()
    target_list.extend(new_items)


def get_field_config(base_dir=None):
    """返回当前生效的字段配置（深拷贝）。

    Returns:
        list: [{"name": str, "keywords": [str]}, ...]
    """
    config = _get_effective_config()
    return [
        {"name": f["name"], "keywords": list(f["keywords"])}
        for f in config
    ]


def _get_field_names():
    """返回当前配置的字段名列表。"""
    config = _get_effective_config()
    return [f["name"] for f in config]


def save_field_config(base_dir, config_list):
    """保存字段配置到 field_config.json，并立即生效。

    Args:
        base_dir: 程序根目录
        config_list: [{"name": str, "keywords": [str]}, ...]
    """
    global _FIELD_CONFIG
    path = os.path.join(base_dir, FIELD_CONFIG_FILENAME)

    # 持久化
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config_list, f, ensure_ascii=False, indent=2)

    # 更新内存缓存
    _FIELD_CONFIG = list(config_list)

    # 同步关键词
    _sync_keywords_from_config()


def reset_field_config(base_dir):
    """重置字段配置为默认值，删除 field_config.json。

    Args:
        base_dir: 程序根目录
    """
    global _FIELD_CONFIG
    path = os.path.join(base_dir, FIELD_CONFIG_FILENAME)

    if os.path.isfile(path):
        os.remove(path)

    _FIELD_CONFIG = []

    # 恢复默认关键词
    _sync_keywords_from_config()


def save_field_mapping(base_dir, header_name, field_type):
    """保存一条字段映射到 settings.json，并立即更新内存映射。

    同时写入 JSON 文件和 CUSTOM_HEADER_MAP 字典，
    确保：① 本次运行立即生效  ② 重启后仍能识别。

    v1.5.0: field_type 支持任意字段名（如 "客户"、"金额"、"日期" 等）。

    Args:
        base_dir: 程序根目录
        header_name: 表头名称（需已 strip 处理）
        field_type: 字段名（如 "客户"、"金额" 等）
    """
    settings_path = os.path.join(base_dir, SETTINGS_FILENAME)

    # 读取现有配置（保留其他字段如 template、statistics 等）
    settings = {}
    if os.path.isfile(settings_path):
        try:
            with open(settings_path, "r", encoding="utf-8") as f:
                settings = json.load(f)
        except (json.JSONDecodeError, OSError):
            settings = {}

    if "field_mapping" not in settings:
        settings["field_mapping"] = {}

    # 写入新映射
    settings["field_mapping"][header_name] = field_type

    # 持久化
    with open(settings_path, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)

    # 立即更新内存映射
    CUSTOM_HEADER_MAP[header_name] = field_type


def get_headers_for_display(ws):
    """读取表头行，返回可供 GUI 下拉框使用的列表。

    Args:
        ws: openpyxl 工作表对象

    Returns:
        list[tuple[int, str]]: [(0-based列号, 表头文本), ...]，跳过空表头
    """
    result = []
    for i, cell in enumerate(ws[1]):
        if cell.value is not None:
            text = str(cell.value).strip()
            if text:
                result.append((i, text))
    return result


# ============================================================
#  v1.3 新增：GUI 服务层接口
#  所有文件操作和识别逻辑封装在此，GUI 只传路径、收结果
# ============================================================

def get_file_headers(file_path):
    """安全打开一个 Excel 文件并返回表头列表。

    GUI 通过此函数获取表头选项，无需直接导入 openpyxl 或 xlrd。
    内部封装了文件打开、工作表定位、表头清洗和异常处理。
    支持 .xlsx / .xlsm / .xls 三种格式。

    Args:
        file_path: Excel 文件的完整路径

    Returns:
        (True, [(0, "客户名称"), (1, "金额"), ...])  — 成功
        (False, "错误描述")                            — 失败
    """
    ext = os.path.splitext(file_path)[1].lower()

    # .xls — 使用 xlrd
    if ext == ".xls":
        try:
            import xlrd
            wb = xlrd.open_workbook(file_path)
            try:
                sheet = wb.sheet_by_index(0)
                if sheet.nrows < 1:
                    return False, "文件表头为空，请在 Excel 第 1 行添加列名"
                row_vals = sheet.row_values(0)
                headers = []
                for idx, val in enumerate(row_vals):
                    text = str(val).strip() if val is not None else ""
                    headers.append((idx, text if text else ""))
                if not headers or all(h[1] == "" for h in headers):
                    return False, "文件表头为空，请在 Excel 第 1 行添加列名"
                return True, headers
            finally:
                wb.release_resources()
        except Exception as e:
            return False, f"无法读取文件：{e}"

    # .xlsx / .xlsm — 使用 openpyxl
    try:
        from openpyxl import load_workbook
        wb = load_workbook(file_path, data_only=True)
        try:
            ws = wb.active
            headers = get_headers_for_display(ws)
        finally:
            wb.close()
        if not headers:
            return False, "文件表头为空，请在 Excel 第 1 行添加列名"
        return True, headers
    except Exception as e:
        return False, f"无法读取文件：{e}"


def get_all_mappings():
    """返回当前所有自定义字段映射的副本。

    供 GUI 映射管理面板展示使用。返回的是浅拷贝，
    修改返回值不影响 CUSTOM_HEADER_MAP。

    Returns:
        dict: {表头名: "customer"|"money", ...}
    """
    return dict(CUSTOM_HEADER_MAP)


def remove_field_mapping(base_dir, header_name):
    """删除一条字段映射（从 settings.json 和内存中同步移除）。

    返回 False 表示映射本来就不存在（调用方可据此决定
    是否提示用户）。

    Args:
        base_dir: 程序根目录
        header_name: 要删除的表头名

    Returns:
        bool: True — 映射存在并已删除；False — 映射不存在
    """
    if header_name not in CUSTOM_HEADER_MAP:
        return False

    # 从内存移除
    del CUSTOM_HEADER_MAP[header_name]

    # 从 settings.json 移除
    settings = _read_settings(base_dir)
    fm = settings.get("field_mapping", {})
    if header_name in fm:
        del fm[header_name]
        settings["field_mapping"] = fm
        _write_settings(base_dir, settings)

    return True
