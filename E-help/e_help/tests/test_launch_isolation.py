"""Packaging diagnostics must never mutate a release or kill unrelated apps."""
from contextlib import redirect_stdout
import importlib.util
import io
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import MagicMock, patch

spec = importlib.util.spec_from_file_location(
    "isolated_launch_tool", Path(__file__).resolve().parents[2] / "tools/test_launch.py")
launch = importlib.util.module_from_spec(spec)
spec.loader.exec_module(launch)


class LaunchIsolationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="ehelper-launch-unit-")
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / "source"
        self.source.mkdir()
        (self.source / launch.EXE_NAME).write_bytes(b"synthetic-exe-never-executed")
        (self.source / "_internal").mkdir()
        (self.source / "_internal/runtime.dll").write_bytes(b"synthetic-runtime")
        log = launch._log_path(self.source)
        log.parent.mkdir(parents=True)
        log.write_text("original-log", encoding="utf-8")
        (self.source / "private.xlsx").write_bytes(b"synthetic-private-bill")
        (self.source / "license.dat").write_bytes(b"synthetic-license")

    def test_only_runtime_copied_original_bytes_untouched_and_no_overwrite(self):
        before = {p.relative_to(self.source): p.read_bytes() for p in self.source.rglob("*") if p.is_file()}
        target = self.folder / "new"
        launch.copy_runtime(self.source, target)
        self.assertEqual({p.name for p in target.iterdir()}, {launch.EXE_NAME, "_internal"})
        self.assertEqual((target / "_internal/runtime.dll").read_bytes(), b"synthetic-runtime")
        with self.assertRaises(FileExistsError):
            launch.copy_runtime(self.source, target)
        self.assertEqual(before, {p.relative_to(self.source): p.read_bytes() for p in self.source.rglob("*") if p.is_file()})

    def test_stop_targets_only_owned_process_and_escalates_only_after_timeout(self):
        proc = MagicMock()
        proc.poll.return_value = None
        proc.wait.side_effect = [subprocess.TimeoutExpired("synthetic", 5), 0]
        with patch.object(launch.subprocess, "run") as command:
            launch._stop_process(proc)
        command.assert_not_called()
        proc.terminate.assert_called_once()
        proc.kill.assert_called_once()
        self.assertEqual(proc.wait.call_count, 2)

    def test_already_exited_process_not_terminated(self):
        proc = MagicMock()
        proc.poll.return_value = 0
        launch._stop_process(proc)
        proc.terminate.assert_not_called()
        proc.kill.assert_not_called()

    def test_existing_data_refuses_launch(self):
        with patch.object(launch.subprocess, "Popen") as popen:
            with self.assertRaises(ValueError):
                launch.launch_and_check(self.source, True)
        popen.assert_not_called()
        self.assertEqual(launch._log_path(self.source).read_text(encoding="utf-8"), "original-log")

    def test_cleanup_on_monitor_exception(self):
        target = self.folder / "new"
        launch.copy_runtime(self.source, target)
        proc = MagicMock()
        proc.poll.return_value = None
        with patch.object(launch.subprocess, "Popen", return_value=proc), patch.object(
                launch.time, "monotonic", side_effect=RuntimeError("synthetic")):
            with self.assertRaises(RuntimeError):
                launch.launch_and_check(target, True)
        proc.terminate.assert_called_once()

    def test_single_exe_timeout_is_not_a_false_pass(self):
        target = self.folder / "new"
        launch.copy_runtime(self.source, target, dependencies=False)
        proc = MagicMock()
        proc.poll.return_value = None
        with patch.object(launch.subprocess, "Popen", return_value=proc), patch.object(launch, "WAIT_SECONDS", 0):
            ok, _ = launch.launch_and_check(target, False)
        self.assertFalse(ok)

    def test_all_scenarios_use_isolated_copies_and_preserve_source(self):
        with patch.object(launch, "ARTIFACT_ROOT", self.folder / "checks"), patch.object(
                launch, "launch_and_check", return_value=(True, "synthetic")) as check, redirect_stdout(io.StringIO()):
            self.assertEqual(launch.run(self.source), 0)
        self.assertEqual(check.call_count, 4)
        for call in check.call_args_list:
            destination = call.args[0]
            self.assertNotEqual(destination, self.source)
            self.assertTrue(destination.is_relative_to(self.folder / "checks"))
            self.assertFalse((destination / launch.DATA_DIR_NAME).exists())
        self.assertEqual(launch._log_path(self.source).read_text(encoding="utf-8"), "original-log")

    def test_help_and_invalid_args_have_no_side_effects(self):
        with patch.object(launch, "run") as run, redirect_stdout(io.StringIO()):
            self.assertEqual(launch.main([]), 2)
            self.assertEqual(launch.main(["--help"]), 0)
            self.assertEqual(launch.main(["--unknown"]), 2)
        run.assert_not_called()


if __name__ == "__main__":
    unittest.main()
