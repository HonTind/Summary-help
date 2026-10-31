"""Named monthly templates and precise aggregation, isolated from legacy settings."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from decimal import Decimal
import os  # Compatibility: existing callers/tests patch monthly.os.replace.

from .analysis import (ReadReport, Record, SheetRules, check_cancel, exact_sum, format_key,
                      ordered_fields, parse_month, read_inputs)

# Compatible public imports; storage has one independent implementation.
from .template_store import (SCHEMA_VERSION, safe_name, validate_month, make_template,
                             validate_template, load_templates, save_template)


@dataclass(frozen=True)
class Aggregate:
    key: str
    amount: Decimal
    records: tuple[Record, ...]
    parts: tuple[str, ...] = ()


@dataclass
class MonthlyResult:
    month: str
    template: dict
    report: ReadReport
    rows: list[Aggregate]

    @property
    def complete(self):
        return self.report.complete

    @property
    def total(self):
        return exact_sum(row.amount for row in self.rows)

    @property
    def config(self):
        return {"mode": "monthly", "month": self.month, "template": deepcopy(self.template),
                "rules": deepcopy(self.report.rules), "key_fields": ordered_fields(self.report.rules),
                "input_files": list(self.report.input_files),
                "date_filter": self.report.rules.get("date_column") is not None,
                "excluded_rows": self.report.excluded_rows}


def group_records(records, cancel=None):
    groups = {}
    for row in records:
        check_cancel(cancel)
        groups.setdefault(row.parts or (row.key,), []).append(row)  # Structured, order-sensitive.
    return groups


def run_monthly(paths, template, month, *, cancel=None, progress=None):
    month = validate_month(month)
    snapshot = validate_template(template)
    report = read_inputs(paths, SheetRules(**snapshot["rules"]), month=month, cancel=cancel, progress=progress)
    groups = group_records(report.records, cancel)
    rows = []
    for parts in sorted(groups):
        check_cancel(cancel)
        records = groups[parts]
        rows.append(Aggregate(format_key(parts), exact_sum(r.amount for r in records), tuple(records), parts))
    return MonthlyResult(month, snapshot, report, rows)
