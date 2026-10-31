"""Customer aggregate / unique-document reconciliation; no first-match guessing."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from decimal import Decimal

from .analysis import (ReadReport, Record, SheetRules, check_cancel, exact_difference, exact_sum,
                      format_key, ordered_fields, parse_amount, read_inputs)
from .monthly import group_records

CATEGORIES = {
    "consistent": "一致", "different": "金额差异", "left_only": "仅左",
    "right_only": "仅右", "ambiguous": "重复歧义",
}


@dataclass(frozen=True)
class Comparison:
    key: str
    category: str
    left_amount: Decimal | None
    right_amount: Decimal | None
    difference: Decimal | None
    left: tuple[Record, ...]
    right: tuple[Record, ...]
    parts: tuple[str, ...] = ()


@dataclass
class ReconciliationResult:
    mode: str
    tolerance: Decimal
    left: ReadReport
    right: ReadReport
    rows: list[Comparison]

    @property
    def complete(self):
        return self.left.complete and self.right.complete and not any(r.category == "ambiguous" for r in self.rows)

    @property
    def all_consistent(self):
        return self.complete and bool(self.rows) and all(r.category == "consistent" for r in self.rows)

    @property
    def counts(self):
        return {category: sum(r.category == category for r in self.rows) for category in CATEGORIES}

    @property
    def config(self):
        return {"mode": self.mode, "tolerance": str(self.tolerance),
                "left_rules": deepcopy(self.left.rules), "right_rules": deepcopy(self.right.rules),
                "left_key_fields": ordered_fields(self.left.rules),
                "right_key_fields": ordered_fields(self.right.rules),
                "left_files": list(self.left.input_files), "right_files": list(self.right.input_files)}


def compare_reports(left: ReadReport, right: ReadReport, mode="customer", tolerance="0", *, cancel=None):
    if mode not in ("customer", "document"):
        raise ValueError("对账模式必须是 customer 或 document")
    tolerance = parse_amount(tolerance)
    if tolerance < 0:
        raise ValueError("容差不能为负数")
    left_groups = group_records(left.records, cancel)
    right_groups = group_records(right.records, cancel)
    rows = []
    for parts in sorted(left_groups.keys() | right_groups.keys()):
        check_cancel(cancel)
        lrows, rrows = tuple(left_groups.get(parts, ())), tuple(right_groups.get(parts, ()))
        if mode == "document" and (len(lrows) > 1 or len(rrows) > 1):
            rows.append(Comparison(format_key(parts), "ambiguous", None, None, None, lrows, rrows, parts))
            continue
        lamount = exact_sum(r.amount for r in lrows) if lrows else None
        ramount = exact_sum(r.amount for r in rrows) if rrows else None
        difference = exact_difference(lamount, ramount) if lrows and rrows else None
        category = "right_only" if not lrows else "left_only" if not rrows else (
            "consistent" if difference.copy_abs() <= tolerance else "different")
        rows.append(Comparison(format_key(parts), category, lamount, ramount, difference, lrows, rrows, parts))
    return ReconciliationResult(mode, tolerance, left, right, rows)


def run_reconciliation(left_paths, right_paths, left_rules: SheetRules, right_rules: SheetRules,
                       mode="customer", tolerance="0", *, cancel=None, progress=None):
    # Validate before opening any files.
    if mode not in ("customer", "document") or parse_amount(tolerance) < 0:
        raise ValueError("对账模式或容差无效")
    if len(left_rules.field_selections()) != len(right_rules.field_selections()):
        raise ValueError("左右两侧键字段数量必须相同，并按顺序一一对应")
    left = read_inputs(left_paths, left_rules, cancel=cancel,
                       progress=(lambda n, total, name: progress("左侧", n, total, name)) if progress else None)
    right = read_inputs(right_paths, right_rules, cancel=cancel,
                        progress=(lambda n, total, name: progress("右侧", n, total, name)) if progress else None)
    return compare_reports(left, right, mode, tolerance, cancel=cancel)
