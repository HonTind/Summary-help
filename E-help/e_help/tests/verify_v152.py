"""
v1.5.2 — 数据目录重构 + 任务记录系统 验证测试

运行:
  cd "Day07(小试牛刀)"
  python -m e_help.tests.verify_v152
"""

import os
import sys
import json
import shutil
import tempfile

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

from e_help import utils, history
from e_help.utils import get_unique_file_path
from e_help.config import APP_NAME, APP_VERSION, APP_FULL_NAME

passed = 0
total = 0


def check(name, condition, detail=""):
    global passed, total
    total += 1
    if condition:
        passed += 1
        print(f"  PASS  {name}")
    else:
        print(f"  FAIL  {name}  {detail}")


# 独立测试环境：把 get_base_dir 指到临时目录，避免污染真实数据目录
TMP = tempfile.mkdtemp(prefix="v152_")
_orig_base_dir = utils.get_base_dir
utils.get_base_dir = lambda: TMP


def restore():
    utils.get_base_dir = _orig_base_dir
    shutil.rmtree(TMP, ignore_errors=True)


print("=" * 55)
print("  v1.5.2 数据目录重构 + 任务记录 验证测试")
print("=" * 55)

# ═══════════════════════════════════════════════════
#  D1-D5: 数据目录子目录
# ═══════════════════════════════════════════════════
print("\n[D] 数据目录子目录结构")

config_dir = utils.get_config_dir()
history_dir = utils.get_history_dir()
logs_dir = utils.get_logs_dir()
data_dir = utils.get_data_dir()

check("D1: config 目录存在", os.path.isdir(config_dir))
check("D2: history 目录存在", os.path.isdir(history_dir))
check("D3: logs 目录存在", os.path.isdir(logs_dir))
check("D4: 各子目录位于数据根目录下",
      os.path.dirname(config_dir) == data_dir
      and os.path.dirname(history_dir) == data_dir
      and os.path.dirname(logs_dir) == data_dir)
check("D5: 数据根目录名为 E-help-data",
      os.path.basename(data_dir) == utils.DATA_DIR_NAME)

# ═══════════════════════════════════════════════════
#  H1-H9: 任务记录
# ═══════════════════════════════════════════════════
print("\n[H] 任务记录系统")

history.clear_records()
check("H1: 清空后为空列表", history.load_history() == [])

rec = {
    "status": "success",
    "input_folder": "/tmp/in",
    "input_files": ["a.xlsx", "b.xlsx"],
    "file_count": 2,
    "output_file": "汇总.xlsx",
    "output_folder": "/tmp/out",
    "fields": ["客户", "金额"],
    "error": "",
}
history.add_record(rec)
loaded = history.load_history()
check("H2: 添加后 1 条记录", len(loaded) == 1)
r0 = loaded[0]
check("H3: 自动补齐 id", bool(r0.get("id")))
check("H4: 自动补齐 time", bool(r0.get("time")))
check("H5: status 正确", r0.get("status") == "success")
check("H6: output_file 正确", r0.get("output_file") == "汇总.xlsx")
check("H7: fields 正确", r0.get("fields") == ["客户", "金额"])

# 失败记录
history.add_record({"status": "failed", "error": "文件被占用"})
loaded = history.load_history()
check("H8: 失败记录存在且带 error", loaded[-1].get("status") == "failed"
      and loaded[-1].get("error") == "文件被占用")

# 删除单条
rid = loaded[0]["id"]
removed = history.delete_record(rid)
loaded = history.load_history()
check("H9: 删除返回 True", removed is True)
check("H10: 删除后剩余 1 条", len(loaded) == 1)
check("H11: 删除不存在记录返回 False", history.delete_record("nope") is False)

# 清空
history.clear_records()
check("H12: 清空后为空", history.load_history() == [])

# ═══════════════════════════════════════════════════
#  H13-H15: 500 条上限
# ═══════════════════════════════════════════════════
print("\n[C] 500 条记录上限")

history.clear_records()
path = history.get_history_path()
recs = [
    {"id": f"r{i:04d}", "time": "2026-08-13 00:00:00",
     "status": "success", "file_count": 1}
    for i in range(500)
]
with open(path, "w", encoding="utf-8") as f:
    json.dump({"version": 1, "records": recs}, f, ensure_ascii=False)

history.add_record({"id": "newest", "status": "failed", "error": "x"})
loaded = history.load_history()
check("H13: 超过上限仍为 500 条", len(loaded) == 500)
check("H14: 最旧记录被删除", not any(r.get("id") == "r0000" for r in loaded))
check("H15: 最新记录在末尾", loaded[-1].get("id") == "newest")

# ═══════════════════════════════════════════════════
#  H16: 损坏的 history.json
# ═══════════════════════════════════════════════════
print("\n[E] 损坏 history.json 容错")

history.clear_records()
with open(path, "w", encoding="utf-8") as f:
    f.write("{ not valid json")
check("H16: 损坏文件返回空列表", history.load_history() == [])

# ═══════════════════════════════════════════════════
#  O1-O3: 输出防覆盖
# ═══════════════════════════════════════════════════
print("\n[O] 输出防覆盖")

out_dir = os.path.join(TMP, "out")
p1 = get_unique_file_path(out_dir, "汇总.xlsx")
open(p1, "w").close()
p2 = get_unique_file_path(out_dir, "汇总.xlsx")
open(p2, "w").close()
p3 = get_unique_file_path(out_dir, "汇总.xlsx")

check("O1: 首次为原文件名", os.path.basename(p1) == "汇总.xlsx")
check("O2: 第二次为 (1)", os.path.basename(p2) == "汇总(1).xlsx")
check("O3: 第三次为 (2)", os.path.basename(p3) == "汇总(2).xlsx")

# ═══════════════════════════════════════════════════
#  V1-V3: 版本统一
# ═══════════════════════════════════════════════════
print("\n[V] 版本统一")

check("V1: APP_VERSION == v1.5.4", APP_VERSION == "v1.5.4")
check("V2: APP_FULL_NAME 含 v1.5.4", "v1.5.4" in APP_FULL_NAME)
check("V3: APP_NAME 正确", APP_NAME == "E-help")

restore()

print("\n" + "=" * 55)
print(f"  {passed}/{total} 通过")
print("=" * 55)
if passed != total:
    sys.exit(1)
