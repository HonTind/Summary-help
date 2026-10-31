"""Workbench-only, read-only Excel input. No legacy configuration or I/O at import."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from decimal import Decimal, InvalidOperation, localcontext
from itertools import zip_longest
import datetime
import json
import math
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import xlrd


class CancelledError(Exception):
    """Cooperative cancellation; callers must discard partial results."""


def check_cancel(cancel=None):
    if cancel is not None and cancel.is_set():
        raise CancelledError("任务已取消，未生成结果")


def _valid_selection(value) -> bool:
    """A field is a non-empty exact header name or a 1-based column number."""
    if type(value) is int and 1 <= value <= 16384:
        return True
    return isinstance(value, str) and bool(value.strip())


MAX_KEY_FIELDS = 8


@dataclass(frozen=True)
class SheetRules:
    key_column: str | int = "客户"
    amount_column: str | int = "金额"
    sheet: str | None = None
    header_row: int = 1
    key_fields: tuple[str | int, ...] | None = None
    date_column: str | int | None = None

    def __post_init__(self):
        if type(self.header_row) is not int or not 1 <= self.header_row <= 1048576:
            raise ValueError("表头行必须是 1 至 1048576 的整数")
        if self.sheet is not None and (not isinstance(self.sheet, str) or not self.sheet.strip()):
            raise ValueError("工作表名称不能为空；默认工作表请用 None")
        for value in (self.key_column, self.amount_column):
            if not _valid_selection(value):
                raise ValueError("字段必须是非空表头名称或从 1 开始的列序号")
        fields = self.key_fields
        if fields is not None:
            # Accept a JSON list/tuple but keep a structured tuple as the sole representation.
            if not isinstance(fields, (list, tuple)):
                raise ValueError("组合键字段必须是列表或元组")
            fields = tuple(fields)
            if not 1 <= len(fields) <= MAX_KEY_FIELDS:
                raise ValueError(f"组合键字段需为 1 至 {MAX_KEY_FIELDS} 项")
            for value in fields:
                if not _valid_selection(value):
                    raise ValueError("组合键字段必须是非空表头名称或从 1 开始的列序号")
            if len(set(fields)) != len(fields):
                raise ValueError("组合键字段不能重复；请检查是否有同一列被填写两次")
            object.__setattr__(self, "key_fields", fields)
        if self.date_column is not None and not _valid_selection(self.date_column):
            raise ValueError("日期字段必须是非空表头名称或从 1 开始的列序号")

    def field_selections(self) -> tuple[str | int, ...]:
        """Ordered key fields; a single key uses the legacy key_column only."""
        return self.key_fields if self.key_fields is not None else (self.key_column,)

    def snapshot(self):
        data = asdict(self)
        if self.key_fields is None:
            data.pop("key_fields")  # Legacy schema-1 layouts stay byte-compatible.
        if self.date_column is None:
            data.pop("date_column")  # Filter off stays byte-compatible with schema 1.
        return data


def parse_month(month):
    """Validate a YYYY-MM label and return (year, month) integers."""
    if not isinstance(month, str) or not re.fullmatch(r"[0-9]{4}-(?:0[1-9]|1[0-2])", month) or month.startswith("0000"):
        raise ValueError("月份必须是有效的 YYYY-MM，例如 2026-10")
    year, mon = month.split("-")
    return int(year), int(mon)


# Excel bases: serial 1 = 1900-01-01 (with the fictitious 1900-02-29) or 1904-01-01.
# In the 1900 base serials 1..59 predate the fictitious leap day and need the 1899-12-31
# offset; serials >= 61 use the standard 1899-12-30 offset. Serial 60 is that fake day.
_1900_EPOCH_EARLY = datetime.date(1899, 12, 31)
_1900_EPOCH_STANDARD = datetime.date(1899, 12, 30)
_MAC_EPOCH_1904 = datetime.date(1904, 1, 1)
_DATE_TEXT = re.compile(r"^([0-9]{4})([-/])([0-9]{2})\2([0-9]{2})$")
_COMPACT_DATE = re.compile(r"^[0-9]{8}$")


def _decode_serial(value, *, datemode) -> datetime.date:
    """Decode an Excel day serial with the workbook base; never silently fix bad dates."""
    number = float(value)
    if not math.isfinite(number):
        raise ValueError("日期序列不是有限数值")
    if datemode == 1:
        # 1904 base: serial 0 is the real 1904-01-01; time fractions keep that same day.
        if number < 0:
            raise ValueError("1904 基准的日期序列不能为负值")
        try:
            return _MAC_EPOCH_1904 + datetime.timedelta(days=number)
        except (OverflowError, ValueError) as exc:
            raise ValueError("日期序列超出可表示范围") from exc
    # 1900 base.
    if number < 1:
        raise ValueError("1900 基准的序列 0 与负值不对应真实日期（序列 0 即虚构的 1900-01-00），拒绝而不静默纠正")
    if math.floor(number) == 60:
        raise ValueError("1900 基准的序列 60 是虚构的 1900-02-29，并非真实日期；拒绝而不静默纠正到 1900-02-28")
    epoch = _1900_EPOCH_STANDARD if number >= 61 else _1900_EPOCH_EARLY
    try:
        return epoch + datetime.timedelta(days=number)
    except (OverflowError, ValueError) as exc:
        raise ValueError("日期序列超出可表示范围") from exc


def _parse_date_text(text) -> datetime.date:
    text = text.strip()
    match = _DATE_TEXT.match(text)
    if match:
        year, month, day = int(match.group(1)), int(match.group(3)), int(match.group(4))
    elif _COMPACT_DATE.match(text):
        year, month, day = int(text[:4]), int(text[4:6]), int(text[6:])
    else:
        raise ValueError("日期文本必须是 YYYY-MM-DD、YYYY/MM/DD 或 YYYYMMDD；"
                         "不接受 01/02/2026 这类有歧义的格式")
    try:
        return datetime.date(year, month, day)
    except ValueError as exc:
        raise ValueError(f"不是真实日期：{text!r}") from exc


def parse_date_value(value, data_type, *, datemode) -> datetime.date:
    """Return the calendar date of one date-field cell, or raise a clear error."""
    if data_type == "e":
        raise ValueError("日期字段是 Excel 错误值，不能作为日期")
    if value is None:
        if data_type == "f":
            raise ValueError("日期字段的公式没有缓存值，请在 Excel 重算并保存后再试")
        raise ValueError("日期字段为空")
    if isinstance(value, bool):
        raise ValueError("日期字段不能是布尔值")
    if isinstance(value, datetime.datetime):
        return value.date()  # Source year/month; no timezone guessing.
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, (int, float)):
        return _decode_serial(value, datemode=datemode)
    if isinstance(value, str):
        return _parse_date_text(value)
    raise ValueError("日期字段类型不支持，必须是真实日期、受限文本格式或 Excel 日期序列")


def _parse_date_cell(value, data_type, serial, *, datemode) -> datetime.date:
    """Decode one date cell, preferring the raw workbook serial when it is numeric.

    openpyxl turns date-formatted numbers (serial 60 in the 1900 base, serial 0 in the
    1904 base, time fractions) into dates/times before analysis can see them, so when
    the sheet XML yielded the original number it decides the result.
    """
    if serial is not None:
        return _decode_serial(serial, datemode=datemode)
    return parse_date_value(value, data_type, datemode=datemode)


def _date_raw_text(value, data_type, serial, datemode) -> str:
    """Trace text for one date cell, keeping the workbook's own number when it matters.

    openpyxl rewrites date-formatted numbers (the fictitious serial 60, the 1904 serial
    0, time fractions) to dates/times before analysis sees them, so when that decoded
    display disagrees with the raw workbook number the number is the traceable source.
    Formula cells keep their text; a cached number never replaces the formula.
    """
    display = "" if value is None else str(value)
    formula = data_type == "f" or (isinstance(value, str) and value.startswith("="))
    if formula or serial is None:
        return display
    try:
        decoded = _decode_serial(serial, datemode=datemode)
    except ValueError:
        decoded = None  # Rejected serials still keep their original number below.
    shown = value.date() if isinstance(value, (datetime.datetime, datetime.date)) else None
    if decoded is not None and decoded == shown:
        return display  # Faithful decode: keep the human-readable source text.
    return str(serial)  # Otherwise the workbook's own number is the traceable source.


def date_trace_text(raw, normalized):
    """Audit text for one date field: the source value and its normalised form."""
    raw = "" if raw is None else str(raw)
    normalized = "" if normalized is None else str(normalized)
    if not raw and not normalized:
        return ""
    if raw.strip() == normalized:
        return normalized or raw
    return f"原 {raw}｜规范 {normalized}"


@dataclass(frozen=True)
class DateFilter:
    year: int
    month: int
    datemode: int  # 0 = 1900 base, 1 = 1904 base

    def contains(self, value) -> bool:
        """Calendar-month membership, i.e. [month start, next month start)."""
        return value.year == self.year and value.month == self.month


@dataclass(frozen=True)
class ExcludedRow:
    """A legal out-of-month row: date verified, key/amount deliberately not checked."""
    source: Source
    raw_date: str
    normalized: str
    raw_key: str = ""       # Genuine source key cells, kept for audit but not verified.
    raw_amount: str = ""    # Genuine source amount cell, never replaced with a fake zero.


def format_key(parts) -> str:
    """Display-only text for a key; matching always uses the structured tuple."""
    parts = tuple(parts)
    if len(parts) == 1:
        return parts[0]
    return json.dumps(list(parts), ensure_ascii=False)


def ordered_fields(rules_snapshot) -> list:
    """Ordered key fields recorded in a rules snapshot, or [] when none were recorded.

    Hand-built reports such as ReadReport(rules={}) carry no field metadata; never
    fabricate a field list for them (that used to raise KeyError on export).
    """
    if not isinstance(rules_snapshot, dict):
        return []
    fields = rules_snapshot.get("key_fields")
    if fields:
        return list(fields)
    if rules_snapshot.get("key_column") is not None:
        return [rules_snapshot["key_column"]]
    return []


@dataclass(frozen=True)
class Source:
    file: str
    sheet: str
    row: int


@dataclass(frozen=True)
class Record:
    key: str
    amount: Decimal
    source: Source
    parts: tuple[str, ...] = ()  # Ordered, stripped key components; the match structure.
    raw_date: str = ""           # Original date cell text (filter on); optional trailing field.
    normalized_date: str = ""    # Verified calendar date as YYYY-MM-DD; optional trailing field.


@dataclass(frozen=True)
class Issue:
    source: Source
    code: str
    message: str
    raw_key: str = ""
    raw_amount: str = ""
    raw_date: str = ""           # Original date cell text; never smuggled into raw_key.
    normalized_date: str = ""    # Verified date when only the key/amount was bad.


@dataclass
class ReadReport:
    records: list[Record] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    scanned_rows: int = 0
    blank_rows: int = 0
    invalid_rows: int = 0
    excluded_rows: int = 0
    files_loaded: int = 0
    input_files: list[str] = field(default_factory=list)
    rules: dict = field(default_factory=dict)
    excluded: list[ExcludedRow] = field(default_factory=list)

    @property
    def complete(self):
        return self.files_loaded > 0 and bool(self.records) and not self.issues


def parse_amount(value) -> Decimal:
    """Plain numeric values only, with bounded precision for safe processing."""
    if isinstance(value, bool) or value is None:
        raise ValueError("金额缺失或为布尔值")
    if not isinstance(value, (str, int, float, Decimal)):
        raise ValueError("金额不是数值或纯数字文本")
    text = str(value).strip()
    if len(text) > 1024 or not re.fullmatch(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", text):
        raise ValueError("金额需为纯数字；不支持空值、非有限值、货币符号或千分位")
    try:
        number = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError("金额无法解析") from exc
    if not number.is_finite() or abs(number.as_tuple().exponent) > 1000 or len(number.as_tuple().digits) > 1000:
        raise ValueError("金额超出安全精度范围（1000 位数字／指数）")
    return number


def exact_sum(values) -> Decimal:
    """Prevent the default 28-digit Decimal context from rounding aggregates."""
    values = tuple(values)
    if not values:
        return Decimal(0)
    places = min(v.as_tuple().exponent for v in values)
    width = max(v.adjusted() for v in values) - places + len(str(len(values))) + 4
    with localcontext() as context:
        context.prec = max(28, width)
        return sum(values, Decimal(0))


def exact_difference(left, right):
    return exact_sum((left, right.copy_negate()))


def _column(headers, selection):
    if type(selection) is int:
        if selection > len(headers):
            raise ValueError(f"列序号 {selection} 超出工作表列数")
        return selection - 1
    indices = [i for i, h in enumerate(headers) if h is not None and str(h).strip() == selection.strip()]
    if len(indices) != 1:
        raise ValueError(f"表头「{selection}」未找到或重名，请指定唯一表头或列序号")
    return indices[0]


def _indices(headers, rules):
    key_indices = tuple(_column(headers, selection) for selection in rules.field_selections())
    amount = _column(headers, rules.amount_column)
    date_idx = _column(headers, rules.date_column) if rules.date_column is not None else None
    if len(set(key_indices)) != len(key_indices):
        raise ValueError("组合键字段不能指向同一列；表头名与列序号指向同列也会被拒绝")
    if amount in key_indices:
        raise ValueError("键字段和金额字段不能是同一列")
    if date_idx is not None and (date_idx in key_indices or date_idx == amount):
        raise ValueError("日期字段不能与键字段或金额字段共用同一列")
    return key_indices, amount, date_idx


def _blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _local_name(tag):
    return tag.rsplit("}", 1)[-1]


def _sheet_part(archive, sheet_name):
    """Resolve a worksheet name to its XML part path through the workbook relationships."""
    try:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, ET.ParseError, OSError):
        return None
    targets = {rel.get("Id"): rel.get("Target")
               for rel in relationships.iter(_PKG_REL_NS + "Relationship")}
    for sheet in workbook.iter(_MAIN_NS + "sheet"):
        if sheet.get("name") == sheet_name:
            target = targets.get(sheet.get(_REL_NS + "id"))
            if not target:
                return None
            if target.startswith("/"):
                return target.lstrip("/")
            if target.startswith("xl/"):
                return target
            return "xl/" + target
    return None


def _xlsx_date_serials(path, sheet_name, date_idx, cancel):
    """Original numeric serials of the date column, read straight from the sheet XML.

    openpyxl decodes date-formatted numbers (the fictitious 1900-02-29 serial 60, the
    1904 serial 0, time fractions) to dates/times before analysis can validate them.
    Keeping the workbook's own <v> lets the selected base decide. Text, boolean, error
    and ISO cells yield no entry and fall back to the parsed value.
    """
    target = get_column_letter(date_idx + 1)
    with zipfile.ZipFile(path) as archive:
        part = _sheet_part(archive, sheet_name)
        if part is None:
            raise ValueError(f"无法定位工作表「{sheet_name}」的 XML，无法按工作簿基准校验日期序列")
        serials = {}
        with archive.open(part) as handle:
            for _, element in ET.iterparse(handle, events=("end",)):
                name = _local_name(element.tag)
                if name == "row":
                    check_cancel(cancel)
                    element.clear()
                    continue
                if name != "c":
                    continue  # Keep <v>/<f> until the enclosing <c> is read.
                reference = element.get("r")
                if reference is not None and element.get("t") in (None, "n"):
                    letters = reference.rstrip("0123456789")
                    row_text = reference[len(letters):]
                    if letters == target and row_text.isdigit():
                        check_cancel(cancel)
                        value = None
                        for child in element:  # A formula keeps its numeric cache here too.
                            if _local_name(child.tag) == "v":
                                value = child.text
                        if value is not None and value.strip():
                            serials[int(row_text)] = value
                element.clear()
        return serials


def _raw_parts(raw, key_indices):
    """Preserve original key cells for the issue table; never substitute blanks."""
    cells = [raw[index] for index in key_indices]
    if len(cells) == 1:
        return str(cells[0])
    return json.dumps(["" if value is None else str(value) for value in cells], ensure_ascii=False)


def _cell_text(raw, index):
    value = raw[index]
    return "" if value is None else str(value)


def _accept(report, source, values, raw, types, indices, date_filter=None, date_serial=None):
    key_indices, amount_idx, date_idx = indices
    required = (*key_indices, amount_idx)
    width = max(required) + 1
    if date_idx is not None:
        width = max(width, date_idx + 1)
    values = list(values) + [None] * max(0, width - len(values))
    raw = list(raw) + [None] * max(0, width - len(raw))
    types = list(types) + [""] * max(0, width - len(types))
    report.scanned_rows += 1
    if all(_blank(v) for v in values):
        report.blank_rows += 1
        return
    raw_date = normalized = ""
    if date_idx is not None:
        # Verify the date first: legal out-of-month rows are skipped without checking key/amount.
        raw_date = _date_raw_text(raw[date_idx], types[date_idx], date_serial,
                                  date_filter.datemode)
        try:
            day = _parse_date_cell(values[date_idx], types[date_idx], date_serial,
                                   datemode=date_filter.datemode)
        except ValueError as exc:
            report.invalid_rows += 1
            # The real date keeps its own columns; never borrowed into the key column.
            report.issues.append(Issue(source, "invalid_date", str(exc),
                                       _raw_parts(raw, key_indices), _cell_text(raw, amount_idx),
                                       raw_date=raw_date))
            return
        normalized = day.isoformat()
        if not date_filter.contains(day):
            report.excluded_rows += 1
            report.excluded.append(ExcludedRow(source, raw_date, normalized,
                                               _raw_parts(raw, key_indices), _cell_text(raw, amount_idx)))
            return
    try:
        for index in required:
            if types[index] == "f" and values[index] is None:
                raise ValueError("必需字段的公式没有缓存值，请在 Excel 重算并保存后再试")
            if types[index] in ("e", "d", "date"):
                raise ValueError("必需字段是 Excel 错误值或日期，不能作为键或金额")
        parts = []
        for position, index in enumerate(key_indices):
            value = values[index]
            if _blank(value) or isinstance(value, bool):
                label = "键字段" if len(key_indices) == 1 else f"第{position + 1}个键字段"
                raise ValueError(f"{label}缺失或为布尔值")
            parts.append(str(value).strip())
        number = parse_amount(values[amount_idx])
        parts = tuple(parts)  # Structured tuple: order-sensitive, no join collisions.
        report.records.append(Record(format_key(parts), number, source, parts, raw_date, normalized))
    except ValueError as exc:
        report.invalid_rows += 1
        report.issues.append(Issue(source, "invalid_row", str(exc),
                                   _raw_parts(raw, key_indices), _cell_text(raw, amount_idx),
                                   raw_date=raw_date, normalized_date=normalized))


def _xlsx(path, rules, report, cancel, target):
    cached = formulas = None
    try:
        cached = load_workbook(path, read_only=True, data_only=True)
        formulas = load_workbook(path, read_only=True, data_only=False)
        name = rules.sheet or cached.active.title
        if name not in cached.sheetnames:
            raise ValueError(f"找不到工作表「{name}」")
        data_ws, raw_ws = cached[name], formulas[name]
        # Ignore inaccurate / absent dimension metadata; scan real XML rows.
        data_ws.reset_dimensions()
        raw_ws.reset_dimensions()
        header_rows = data_ws.iter_rows(min_row=rules.header_row, max_row=rules.header_row, values_only=True)
        try:
            header = next(header_rows, ())
        finally:
            header_rows.close()
        if not header or all(_blank(value) for value in header):
            raise ValueError("表头行缺失或全空")
        indices = _indices(header, rules)
        datemode = 1 if getattr(cached, "epoch", None) is not None and cached.epoch.year == 1904 else 0
        date_filter = DateFilter(target[0], target[1], datemode) if indices[2] is not None else None
        date_serials = (_xlsx_date_serials(path, name, indices[2], cancel)
                        if indices[2] is not None else None)
        first = rules.header_row + 1
        data_rows = data_ws.iter_rows(min_row=first)
        raw_rows = raw_ws.iter_rows(min_row=first)
        for row_number, (cells, raw_cells) in enumerate(zip_longest(data_rows, raw_rows), first):
            check_cancel(cancel)
            if cells is None or raw_cells is None:
                raise ValueError("缓存工作表与原始工作表行数不一致")
            values = [c.value for c in cells]
            raw = [c.value for c in raw_cells]
            types = [("e" if getattr(cell, "data_type", "") == "e" else
                      "d" if getattr(cell, "data_type", "") == "d" else getattr(original, "data_type", ""))
                     for cell, original in zip_longest(cells, raw_cells)]
            date_serial = date_serials.get(row_number) if date_serials else None
            _accept(report, Source(str(path), name, row_number), values, raw, types, indices,
                    date_filter, date_serial)
        report.files_loaded += 1
    finally:
        if cached is not None:
            cached.close()
        if formulas is not None:
            formulas.close()


def _xls(path, rules, report, cancel, target):
    book = xlrd.open_workbook(str(path), on_demand=True)
    try:
        sheet = book.sheet_by_name(rules.sheet) if rules.sheet else book.sheet_by_index(0)
        if rules.header_row > sheet.nrows:
            raise ValueError("表头行超出工作表范围")
        indices = _indices(sheet.row_values(rules.header_row - 1), rules)
        date_filter = DateFilter(target[0], target[1], datemode=book.datemode) if indices[2] is not None else None
        for index in range(rules.header_row, sheet.nrows):
            check_cancel(cancel)
            row = sheet.row(index)
            values = [bool(c.value) if c.ctype == xlrd.XL_CELL_BOOLEAN else c.value for c in row]
            types = ["e" if c.ctype == xlrd.XL_CELL_ERROR else
                     "date" if c.ctype == xlrd.XL_CELL_DATE else "" for c in row]
            _accept(report, Source(str(path), sheet.name, index + 1), values, values, types, indices, date_filter)
        report.files_loaded += 1
    finally:
        book.release_resources()


def read_inputs(paths, rules: SheetRules, *, month=None, cancel=None, progress=None) -> ReadReport:
    """Read only explicitly selected files; failures remain visible in the report."""
    if not isinstance(rules, SheetRules):
        raise ValueError("必须提供 SheetRules")
    if rules.date_column is not None:
        if month is None:
            raise ValueError("启用日期筛选时必须提供月份（YYYY-MM）")
        target = parse_month(month)
    else:
        target = None
    files = [Path(p).resolve() for p in paths]
    report = ReadReport(input_files=[str(p) for p in files], rules=rules.snapshot())
    seen = set()
    for index, path in enumerate(files, 1):
        check_cancel(cancel)
        source = Source(str(path), rules.sheet or "", rules.header_row)
        identity = str(path).casefold()
        if identity in seen:
            report.issues.append(Issue(source, "duplicate_file", "同一文件重复选择，未重复计入"))
            continue
        seen.add(identity)
        try:
            if path.suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
                raise ValueError("仅支持 .xlsx / .xlsm / .xls")
            before = len(report.records)
            before_excluded = report.excluded_rows
            (_xls if path.suffix.lower() == ".xls" else _xlsx)(path, rules, report, cancel, target)
            if len(report.records) == before and report.excluded_rows == before_excluded:
                message = "该文件没有本月有效数据行" if rules.date_column is not None else "该文件没有有效数据行"
                report.issues.append(Issue(source, "no_valid_rows", message))
        except CancelledError:
            raise
        except Exception as exc:
            report.issues.append(Issue(source, "file_error", f"{type(exc).__name__}: {exc}"))
        if progress:
            progress(index, len(files), path.name)
    if not files:
        report.issues.append(Issue(Source("", "", 0), "no_input", "请先选择输入文件"))
    return report
