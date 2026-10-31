"""
任务记录模块 — history.json 管理

职责：
    - 管理本地 history.json 的读写
    - 添加 / 删除 / 清空任务记录
    - 自动截断到 MAX_RECORDS 条（删除最旧记录）

原则：
    - 纯本地，不联网
    - 读取异常不崩溃，统一返回空列表
    - 删除记录只影响 history.json，绝不触碰实际 Excel 文件

接口：
    add_record(record) -> None
    delete_record(record_id) -> bool
    clear_records() -> None
    load_history() -> list
"""

import os
import json
import datetime
import tempfile
import threading
import uuid

from .utils import get_history_dir

HISTORY_FILENAME = "history.json"
HISTORY_VERSION = 1
MAX_RECORDS = 500
_LOCK = threading.RLock()
TASK_TYPES = {"regular": "常规汇总", "monthly": "月度汇总",
              "customer": "客户总额对账", "document": "单据逐笔对账"}
STATUS_LABELS = {"success": "已导出", "previewed": "预览完成", "incomplete": "预览不完整",
                 "failed": "失败", "cancelled": "已取消"}
PHASE_LABELS = {"preview": "预览", "export": "导出", "summary": "汇总"}


def get_history_path():
    """返回 history.json 的完整路径。"""
    return os.path.join(get_history_dir(), HISTORY_FILENAME)


def _now_str():
    """返回当前时间字符串（YYYY-MM-DD HH:MM:SS）。"""
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _new_id():
    """生成任务记录的唯一 ID（微秒时间戳 + 序号前缀）。

    正常使用场景下同一微秒内不会连续生成两条记录，足以保证唯一性。
    """
    now = datetime.datetime.now()
    return now.strftime("%Y%m%d%H%M%S%f")


def _read_store():
    """Read without rewriting. Invalid/unknown files must never become an empty write."""
    try:
        path = get_history_path()
        if not os.path.lexists(path):
            return {"version": HISTORY_VERSION, "records": []}, ""
        with open(path, "r", encoding="utf-8-sig") as handle:
            data = json.load(handle)
        if not isinstance(data, dict) or not isinstance(data.get("records"), list):
            return None, "记录文件损坏或结构无效，原文件已保留。"
        version = data.get("version", HISTORY_VERSION)
        if type(version) is not int or version != HISTORY_VERSION:
            return None, "记录文件版本不支持，原文件已保留。"
        if any(not isinstance(r, dict) or not isinstance(r.get("id"), str) or not r["id"] for r in data["records"]):
            return None, "记录中存在无效项目，原文件已保留。"
        if len({r["id"] for r in data["records"]}) != len(data["records"]):
            return None, "记录中存在重复标识，原文件已保留。"
        return data, ""
    except (OSError, ValueError, UnicodeError):
        return None, "记录文件无法读取，原文件已保留。"


def read_history():
    """Return (records, safe warning); no original error text or private file paths."""
    with _LOCK:
        data, warning = _read_store()
        return (list(data["records"]) if data is not None else []), warning


def load_history():
    return read_history()[0]


def _save(records):
    """Atomic replacement in the same directory; preserve unknown root metadata."""
    temporary = None
    with _LOCK:
        try:
            existing, warning = _read_store()
            if warning:
                return False
            data = dict(existing, version=HISTORY_VERSION, records=records)
            path = get_history_path()
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=os.path.dirname(path),
                                             prefix=".history-", suffix=".tmp", delete=False) as handle:
                temporary = handle.name
                json.dump(data, handle, ensure_ascii=False, indent=2, allow_nan=False)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, path)
            temporary = None
            return True
        except (OSError, ValueError, TypeError):
            return False
        finally:
            if temporary is not None:
                try:
                    os.unlink(temporary)  # Only the temp file created by this invocation.
                except OSError:
                    pass


def add_record(record):
    """添加一条任务记录，自动截断到 MAX_RECORDS 条。

    Args:
        record: dict，包含 id/time/status/... 等字段
    """
    append_record(record)  # Legacy API keeps its None return and caller ID/time completion.


def append_record(record):
    if not isinstance(record, dict):
        return False
    with _LOCK:
        records, warning = read_history()
        if warning:
            return False
        if not record.get("id"):
            record["id"] = _new_id()
        if not record.get("time"):
            record["time"] = _now_str()
        if not isinstance(record["id"], str):
            return False
        if any(r["id"] == record["id"] for r in records):
            return False
        return _save([*records, dict(record)][-MAX_RECORDS:])


def save_task(task_type, status, *, phase, file_count=0, output_path="", **counts):
    """New UI records only whitelisted metadata, never inputs/keys/amounts/error text."""
    if (not isinstance(task_type, str) or not isinstance(status, str) or not isinstance(phase, str)
            or task_type not in TASK_TYPES or status not in STATUS_LABELS or phase not in PHASE_LABELS):
        return False
    if ((task_type == "regular") != (phase == "summary")
            or (status in ("previewed", "incomplete") and phase != "preview")
            or (status == "success" and phase == "preview")):
        return False
    def number(value):
        return value if type(value) is int and value >= 0 else 0
    record = {"id": uuid.uuid4().hex, "task_type": task_type, "status": status, "phase": phase, "file_count": number(file_count)}
    for name in ("valid_row_count", "result_count", "issue_count", "excluded_count"):
        if name in counts:
            record[name] = number(counts[name])
    if type(counts.get("complete")) is bool:
        record["complete"] = counts["complete"]
    record["output_file"] = record["output_folder"] = ""
    if status == "success" and output_path:
        record["output_file"] = os.path.basename(os.fspath(output_path))
        record["output_folder"] = os.path.dirname(os.fspath(output_path))
    return append_record(record)


def task_type(record):
    value = record.get("task_type", "regular")
    return value if isinstance(value, str) else "unknown"


def type_label(record):
    return TASK_TYPES.get(task_type(record), "其他任务")


def status_label(record):
    value = record.get("status")
    status = STATUS_LABELS.get(value, "未知状态") if isinstance(value, str) else "未知状态"
    if record.get("status") == "success" and record.get("complete") is False:
        status += "（不完整）"
    if value in ("failed", "cancelled") and isinstance(record.get("phase"), str) and record["phase"] in PHASE_LABELS:
        status += " · " + PHASE_LABELS[record["phase"]]
    return status


def filter_records(records, kind=None):
    return [r for r in records if kind is None or task_type(r) == kind]


def output_path(record):
    """Return a local Excel path only; no URL/UNC launch from edited record files."""
    if record.get("status") != "success":
        return ""
    folder, name = record.get("output_folder", ""), record.get("output_file", "")
    if not isinstance(folder, str) or not isinstance(name, str) or not folder or not name:
        return ""
    path = os.path.join(folder, name)
    if path.startswith(("\\\\", "//")) or "://" in path or "\x00" in path:
        return ""
    if os.path.splitext(path)[1].lower() not in (".xlsx", ".xls", ".xlsm"):
        return ""
    return path


def delete_record(record_id):
    """删除指定 id 的记录。返回是否确实删除了记录。

    仅删除 history.json 中的记录，绝不触碰实际文件。
    """
    with _LOCK:
        records, warning = read_history()
        if warning:
            return False
        new_records = [r for r in records if r.get("id") != record_id]
        return len(new_records) < len(records) and _save(new_records)


def clear_records():
    """清空所有任务记录（仅清空记录，不删除任何实际文件）。"""
    _save([])


def clear_records_safe():
    return _save([])
