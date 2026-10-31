import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from e_help import utils


class BrandPathsTests(unittest.TestCase):
    def test_new_install_uses_new_data_name(self):
        with tempfile.TemporaryDirectory() as root, patch.object(utils, "get_base_dir", return_value=root):
            self.assertEqual(Path(utils.get_data_dir()), Path(root) / "E-help-data")

    def test_existing_data_remains_accessible(self):
        with tempfile.TemporaryDirectory() as root, patch.object(utils, "get_base_dir", return_value=root):
            legacy = Path(root) / utils.LEGACY_DATA_DIR_NAME
            legacy.mkdir()
            marker = legacy / "history.json"
            marker.write_text("[]", encoding="utf-8")
            self.assertEqual(Path(utils.get_data_dir()), legacy)
            self.assertEqual(marker.read_text(encoding="utf-8"), "[]")
            current = Path(root) / utils.DATA_DIR_NAME
            current.mkdir()
            self.assertEqual(Path(utils.get_data_dir()), current)
