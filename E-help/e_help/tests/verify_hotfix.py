"""
v1.3.2 Hotfix — start_gui.py 回归验证

运行:
  cd "Day07(小试牛刀)"
  python -m e_help.tests.verify_hotfix
"""

import os, sys, shutil

THIS_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(THIS_DIR)
PARENT_DIR = os.path.dirname(PROJECT_ROOT)
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)

TEST_DIR = os.path.join(PROJECT_ROOT, "tests", "_hotfix_test")
if os.path.isdir(TEST_DIR):
    shutil.rmtree(TEST_DIR)
os.makedirs(TEST_DIR)

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

print("=" * 55)
print("  v1.3.2 Hotfix - start_gui.py regression")
print("=" * 55)

# T5: Regression
print("\n[T5] Regression - core functions")

from e_help.field_mapper import load_settings, save_field_mapping, get_all_mappings
import e_help.config as config_module
_orig = dict(config_module.CUSTOM_HEADER_MAP)
config_module.CUSTOM_HEADER_MAP.clear()
config_module.CUSTOM_HEADER_MAP.update(_orig)

save_field_mapping(TEST_DIR, "test_regression", "customer")
check("T5a: field_mapper save ok",
      config_module.CUSTOM_HEADER_MAP.get("test_regression") == "customer")

from e_help.gui import main
check("T5b: gui.main importable", callable(main))

from e_help.gui import _WorkerThread
check("T5c: _WorkerThread importable", True)

load_settings(TEST_DIR)
check("T5d: load_settings no crash", True)

# T6: start_gui.py content
print("\n[T6] start_gui.py content check")
start_py = os.path.join(PARENT_DIR, "start_gui.py")
with open(start_py, "r", encoding="utf-8") as f:
    content = f.read()
check("T6a: imports gui.main",
      "from e_help.gui import main" in content)
check("T6b: calls main()",
      "main()" in content)
check("T6c: no direct SummaryApp",
      "SummaryApp(" not in content)

# cleanup
config_module.CUSTOM_HEADER_MAP.clear()
config_module.CUSTOM_HEADER_MAP.update(_orig)
shutil.rmtree(TEST_DIR)

print(f"\n{'=' * 55}")
print(f"  {passed}/{total} PASS")
print(f"{'=' * 55}")
