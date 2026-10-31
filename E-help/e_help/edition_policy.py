"""Unified offline edition: no activation, device binding or export quota."""
from contextlib import contextmanager
from functools import wraps

class FreePolicy:
    @contextmanager
    def export(self):
        # Keep the shared export boundary without touching any entitlement state.
        yield


def get_policy():
    return FreePolicy()


def edition_status():
    from .config import APP_FULL_NAME
    return f"{APP_FULL_NAME}\nExcel 汇总、月度处理与账单对账工具。\n业务数据在本机处理。"


def guarded_export(function):
    @wraps(function)
    def wrapped(*args, **kwargs):
        with get_policy().export():
            return function(*args, **kwargs)
    return wrapped
