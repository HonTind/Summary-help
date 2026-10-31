"""Validated, atomic monthly template storage; no GUI or aggregation."""
from __future__ import annotations

from copy import deepcopy
import json
import os
from pathlib import Path
import re
import tempfile

from .analysis import SheetRules, parse_month

SCHEMA_VERSION = 1


def safe_name(value, *, max_length=80):
    if not isinstance(value, str) or not value.strip() or len(value) > max_length:
        raise ValueError(f"名称必须是 1 至 {max_length} 字符")
    if value != value.strip() or re.search(r'[<>:"/\\|?*\x00-\x1f]', value) or value.endswith("."):
        raise ValueError("名称不能含路径、控制字符或 Windows 非法字符")
    if value in (".", "..") or value.split(".")[0].upper() in (
        "CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))
    ):
        raise ValueError("名称不能是系统保留名称")
    return value


def validate_month(month):
    parse_month(month)
    return month


def make_template(name, rules: SheetRules, output_prefix="月度汇总"):
    return validate_template({
        "schema_version": SCHEMA_VERSION, "name": safe_name(name),
        "rules": rules.snapshot(), "output_prefix": safe_name(output_prefix),
    })


def validate_template(template):
    expected = {"schema_version", "name", "rules", "output_prefix"}
    if not isinstance(template, dict) or set(template) != expected:
        raise ValueError("模板必须只含版本、名称、字段规则和输出前缀，不能保存账单内容或旧文件名单")
    if type(template["schema_version"]) is not int or template["schema_version"] != SCHEMA_VERSION:
        raise ValueError("不支持的模板版本")
    safe_name(template["name"])
    safe_name(template["output_prefix"])
    rules = template["rules"]
    legacy = {"key_column", "amount_column", "sheet", "header_row"}
    # The optional date rule only appears when filtering is on, so schema-1 layouts stay valid.
    allowed = (legacy, legacy | {"key_fields"},
               legacy | {"date_column"}, legacy | {"key_fields", "date_column"})
    if not isinstance(rules, dict) or set(rules) not in allowed:
        raise ValueError("模板字段规则无效")
    SheetRules(**rules)
    return deepcopy(template)


def load_templates(path):
    path = Path(path)
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if not isinstance(data, dict) or set(data) != {"schema_version", "templates"}:
        raise ValueError("模板文件结构无效；保留原文件")
    if type(data["schema_version"]) is not int or data["schema_version"] != SCHEMA_VERSION:
        raise ValueError("不支持的模板文件版本；保留原文件")
    if not isinstance(data["templates"], list):
        raise ValueError("模板列表无效")
    result = {}
    for entry in data["templates"]:
        template = validate_template(entry)
        name = template["name"]
        if name in result:
            raise ValueError("模板文件存在重名；保留原文件")
        result[name] = template
    return result


def save_template(path, template, *, allow_update=False):
    """Atomic local save; the caller must obtain approval before an existing-name update."""
    template = validate_template(template)
    path = Path(path)
    templates = load_templates(path)  # Do not reset broken JSON or unsupported versions.
    if template["name"] in templates and not allow_update:
        raise FileExistsError("同名模板已存在，需要用户明确批准更新")
    templates[template["name"]] = template
    payload = {"schema_version": SCHEMA_VERSION, "templates": list(templates.values())}
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile("w", dir=path.parent, prefix=".monthly-", suffix=".tmp",
                                         encoding="utf-8", delete=False) as handle:
            temporary = Path(handle.name)
            json.dump(payload, handle, ensure_ascii=False, indent=2, allow_nan=False)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
        temporary = None
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)  # Only our newly-created temporary file.
    return deepcopy(template)
