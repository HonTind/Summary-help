"""Free edition regression: no signed-license fixtures or real user state."""
import importlib.util
import io
import os
from pathlib import Path
import subprocess
import sys
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from e_help import edition_policy as policy
from e_help.excel_writer import generate_summary


class FreeTests(unittest.TestCase):
    def test_more_than_ten_exports_without_local_profile(self):
        with patch.dict(os.environ, {}, clear=True):
            for _ in range(12):
                output = io.BytesIO()
                generate_summary([], {}, output)
                self.assertTrue(output.getvalue().startswith(b"PK"))

    def test_status_and_policy_do_not_access_files(self):
        with patch("builtins.open", side_effect=AssertionError("no files")), \
             patch.object(Path, "open", side_effect=AssertionError("no files")):
            self.assertIn("E-help", policy.edition_status())
            for _ in range(12):
                with policy.get_policy().export():
                    pass
        for name in ("machine_code", "install_license", "license_path", "PaidPolicy"):
            self.assertFalse(hasattr(policy, name))

    def test_export_errors_propagate_and_do_not_consume_quota(self):
        @policy.guarded_export
        def fail():
            raise OSError("synthetic failure")
        with self.assertRaisesRegex(OSError, "synthetic failure"):
            fail()
        self.test_more_than_ten_exports_without_local_profile()

    def test_runtime_has_no_license_dependencies(self):
        script = (
            "import sys; before = set(sys.modules); from e_help import edition_policy as p; "
            "assert not any(n == 'winreg' or n.startswith('cryptography') "
            "or n.endswith('paid_public_key') for n in set(sys.modules) - before); "
            "assert 'E-help' in p.edition_status()"
        )
        subprocess.run([sys.executable, "-B", "-c", script], cwd=ROOT, check=True)


def main():
    if len(sys.argv) == 1:
        return max(subprocess.call([sys.executable, "-B", __file__, group])
                   for group in ("core", "test_qt_ui", "test_qt_production"))
    loader = unittest.TestLoader()
    suite = unittest.TestSuite()
    if sys.argv[1] == "core":
        suite.addTests(loader.discover(str(ROOT / "e_help" / "tests"), pattern="test_*.py"))
        suite.addTests(loader.loadTestsFromTestCase(FreeTests))
    else:
        name = sys.argv[1]
        spec = importlib.util.spec_from_file_location(name, ROOT / "tools" / f"{name}.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        suite.addTests(loader.loadTestsFromModule(module))
    result = unittest.TextTestRunner(verbosity=1).run(suite)
    return 0 if result.wasSuccessful() else 1


if __name__ == "__main__":
    raise SystemExit(main())
