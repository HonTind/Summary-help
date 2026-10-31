import os
import sys

# 数据目录名称
DATA_DIR_NAME = "E-help-data"
LEGACY_DATA_DIR_NAME = "Excel智能汇总助手数据"

# 子目录名称
CONFIG_DIR_NAME = "config"
HISTORY_DIR_NAME = "history"
LOGS_DIR_NAME = "logs"


def get_base_dir():
    """获取程序所在目录（兼容 PyInstaller 打包）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_data_dir():
    """获取程序数据根目录，不存在时自动创建。

    Returns:
        str: 数据目录绝对路径
    """
    base_dir = get_base_dir()
    data_dir = os.path.join(base_dir, DATA_DIR_NAME)
    legacy_dir = os.path.join(base_dir, LEGACY_DATA_DIR_NAME)
    if not os.path.exists(data_dir) and os.path.isdir(legacy_dir):
        data_dir = legacy_dir
    os.makedirs(data_dir, exist_ok=True)
    return data_dir


def _get_subdir(name):
    """返回数据根目录下的子目录，不存在时自动创建。"""
    sub = os.path.join(get_data_dir(), name)
    os.makedirs(sub, exist_ok=True)
    return sub


def get_config_dir():
    """返回配置目录（field_config.json / settings.json）。"""
    return _get_subdir(CONFIG_DIR_NAME)


def get_history_dir():
    """返回任务记录目录（history.json）。"""
    return _get_subdir(HISTORY_DIR_NAME)


def get_logs_dir():
    """返回日志目录（startup.log / error.log）。"""
    return _get_subdir(LOGS_DIR_NAME)


def get_unique_file_path(directory, filename):
    """生成不重复的文件路径，防止覆盖已有文件。

    规则：
        - 文件名不存在 → 原样返回
        - 已存在 → 添加编号 (1), (2), (3) ...

    Args:
        directory: 目标目录
        filename: 期望的文件名（如 "汇总.xlsx"）

    Returns:
        str: 唯一的文件完整路径

    Example:
        >>> get_unique_file_path("/out", "汇总.xlsx")
        '/out/汇总.xlsx'          # 不存在
        '/out/汇总(1).xlsx'       # 汇总.xlsx 已存在
        '/out/汇总(2).xlsx'       # 汇总(1).xlsx 也已存在
    """
    if not os.path.isdir(directory):
        os.makedirs(directory, exist_ok=True)

    base, ext = os.path.splitext(filename)
    candidate = os.path.join(directory, filename)

    if not os.path.exists(candidate):
        return candidate

    counter = 1
    while True:
        candidate = os.path.join(directory, f"{base}({counter}){ext}")
        if not os.path.exists(candidate):
            return candidate
        counter += 1
