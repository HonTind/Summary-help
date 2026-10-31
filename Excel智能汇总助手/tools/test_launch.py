# -*- coding: utf-8 -*-
"""Offline onedir launch checks using isolated, newly-created runtime copies.

Only the EXE and _internal are copied; existing logs, configuration and business
files beside the source EXE are never read, changed or removed. Test copies and
their logs remain beneath 本地产物/Excel智能汇总助手/build/ for diagnosis.
"""
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ARTIFACT_ROOT = PROJECT_ROOT.parent / "本地产物" / "Excel智能汇总助手" / "build"
EXE_NAME = "E助手.exe"
DATA_DIR_NAME = "Excel智能汇总助手数据"
WAIT_SECONDS = 12

USAGE = """用法：
    python tools/test_launch.py <onedir发行目录>

只在新的隔离副本中启动 EXE；不在原发行目录运行，不删除原日志，不结束其他同名程序。
副本仅包含 E助手.exe 和 _internal，放在仓库根 本地产物/Excel智能汇总助手/build/ 下。
测试结束保留副本与日志供检查；无参数或 --help 不启动程序或创建目录。

退出码：0 = 全部通过，1 = 测试失败，2 = 用法或发行目录不合法。"""


def _log_path(dist_dir):
    return Path(dist_dir) / DATA_DIR_NAME / "logs" / "startup.log"


def copy_runtime(source, destination, *, dependencies=True):
    """Copy runtime components only; refuse to merge with an existing target."""
    source, destination = Path(source), Path(destination)
    if not (source / EXE_NAME).is_file():
        raise ValueError("发行目录缺少 E助手.exe")
    if dependencies and not (source / "_internal").is_dir():
        raise ValueError("只支持含 _internal 的 onedir 发行目录")
    destination.mkdir(parents=True, exist_ok=False)
    shutil.copy2(source / EXE_NAME, destination / EXE_NAME)
    if dependencies:
        shutil.copytree(
            source / "_internal", destination / "_internal",
            ignore=shutil.ignore_patterns(DATA_DIR_NAME, "Excel汇总结果", "license.dat",
                                          "*.log", ".env", ".env.*", ".claude", ".codex",
                                          ".deepseek", "*.xlsx", "*.xls", "*.xlsm"),
        )


def _stop_process(proc):
    """Release only the process created by this invocation, never a name match."""
    if proc.poll() is None:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            proc.kill()
            proc.wait(timeout=5)
    else:
        proc.wait(timeout=5)


def launch_and_check(dist_dir, expect_mainloop):
    """Check a fresh isolated copy; an existing data directory is never reused."""
    dist_dir = Path(dist_dir)
    if (dist_dir / DATA_DIR_NAME).exists():
        raise ValueError("启动测试必须使用全新副本，不能复用已有运行数据")
    exe, log = dist_dir / EXE_NAME, _log_path(dist_dir)
    if not exe.is_file():
        return False, "隔离副本中缺少 EXE"
    proc = subprocess.Popen([str(exe)], cwd=str(dist_dir))
    content = ""
    try:
        deadline = time.monotonic() + WAIT_SECONDS
        while True:
            if log.is_file():
                content = log.read_text(encoding="utf-8", errors="replace")
            reached = "mainloop begin" in content
            exited = proc.poll()
            if reached or exited is not None or time.monotonic() >= deadline:
                break
            time.sleep(0.1)
        # A broken/missing-DLL bundle must actually exit unsuccessfully, not hang.
        ok = (reached and exited is None) if expect_mainloop else (
            not reached and exited is not None and exited != 0)
        return ok, _summarize(content)
    finally:
        _stop_process(proc)


def _summarize(content):
    markers = ("start", "create SummaryApp begin", "create SummaryApp end", "mainloop begin")
    found = [marker for marker in markers if marker in content]
    if not content:
        return "（无 startup.log）"
    return "日志节点：" + " → ".join(found) if found else "（日志为空或仅部分写入）"


def run(dist_dir):
    source = Path(dist_dir).resolve()
    if not (source / EXE_NAME).is_file() or not (source / "_internal").is_dir():
        print("错误：请指定包含 E助手.exe 与 _internal 的 onedir 发行目录。")
        return 2
    ARTIFACT_ROOT.mkdir(parents=True, exist_ok=True)
    # Unique session: no overwrite and no recursive cleanup of historical artifacts.
    session = Path(tempfile.mkdtemp(prefix="launch-check-", dir=ARTIFACT_ROOT))
    print(f"隔离启动测试；原发行目录保持只读。测试副本：{session}")
    results = []
    for label, name, dependencies in (
        ("A 干净副本启动", "baseline", True),
        ("B 含空格路径", "path with spaces/E助手", True),
        ("C 中文路径", "软件测试/E助手", True),
        ("D 缺失依赖的单 EXE", "exe-only", False),
    ):
        destination = session / name
        try:
            copy_runtime(source, destination, dependencies=dependencies)
            ok, evidence = launch_and_check(destination, expect_mainloop=dependencies)
        except (OSError, ValueError, subprocess.SubprocessError) as exc:
            ok, evidence = False, f"测试无法完成：{type(exc).__name__}"
        results.append(ok)
        print(f"{'PASS' if ok else 'FAIL'} {label}：{evidence}")
    print(f"副本与日志已保留：{session}")
    return 0 if all(results) else 1


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if not argv:
        print(USAGE)
        return 2
    if argv in (["-h"], ["--help"]):
        print(USAGE)
        return 0
    if len(argv) != 1 or argv[0].startswith("-"):
        print(USAGE)
        return 2
    return run(argv[0])


if __name__ == "__main__":
    sys.exit(main())
