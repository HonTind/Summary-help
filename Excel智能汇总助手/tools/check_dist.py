# -*- coding: utf-8 -*-
r"""
发行完整性检查（Packaging Fix，v1.5.4）

用于 PyInstaller onedir 打包完成后，校验最终发行目录是否完整。

onedir 发行物不是单个 EXE，而是（发行目录由命令行参数指定）：
    <发行目录>/
    ├── E助手.exe
    └── _internal/
        ├── python312.dll
        ├── base_library.zip
        └── ... 其他运行依赖

本脚本检查以下关键文件是否存在且非空，任何一项缺失都会以非零退出码失败：

    - E助手.exe          引导程序
    - _internal/python312.dll        Python 解释器 DLL
    - _internal/base_library.zip     Python 标准库
    - _internal/VCRUNTIME140.dll     MSVC 运行库
    - _internal/VCRUNTIME140_1.dll   MSVC 运行库
    - _internal/_tkinter.pyd         Tk 界面扩展
    - _internal/tcl86t.dll           Tcl 运行时
    - _internal/tk86t.dll            Tk 运行时
    - _internal/zlib1.dll            zlib（openpyxl 依赖）

用法：
    python tools/check_dist.py <发行目录> [--qt]

    发行目录必填，没有默认值：发行产物已集中存放，脚本不再猜测目录位置。
    集中存放后的位置示例（在项目目录中执行，release 名替换为实际值）：

        python tools/check_dist.py ..\本地产物\Excel智能汇总助手\dist\<release>\E助手

    也可以显式传入其他发行目录，例如迁移前的旧 dist 目录。

退出码：
    0 = 完整性检查通过
    1 = 缺失关键运行文件
    2 = 用法错误（缺少发行目录或参数不合法），不会读取任何发行目录
"""

import os
import sys

# 关键运行文件：缺一不可
REQUIRED_FILES = [
    ("E助手.exe", "引导程序 EXE"),
    ("_internal/python312.dll", "Python 解释器 DLL"),
    ("_internal/base_library.zip", "Python 标准库"),
    ("_internal/VCRUNTIME140.dll", "MSVC 运行库"),
    ("_internal/VCRUNTIME140_1.dll", "MSVC 运行库"),
    ("_internal/_tkinter.pyd", "Tk 界面扩展"),
    ("_internal/tcl86t.dll", "Tcl 运行时"),
    ("_internal/tk86t.dll", "Tk 运行时"),
    ("_internal/zlib1.dll", "zlib（openpyxl 依赖）"),
]

# 次要运行库：丢失时仅告警，不致命（不同打包环境可能名称略有差异）
OPTIONAL_FILES = [
    ("_internal/libcrypto-3.dll", "OpenSSL 加密库"),
    ("_internal/libssl-3.dll", "OpenSSL SSL 库"),
    ("_internal/libffi-8.dll", "libffi"),
    ("_internal/_ssl.pyd", "SSL 扩展"),
    ("_internal/_hashlib.pyd", "哈希扩展"),
]

QT_REQUIRED_FILES = [
    ("_internal/PySide6/QtCore.pyd", "Qt Core Python binding"),
    ("_internal/PySide6/QtGui.pyd", "Qt GUI Python binding"),
    ("_internal/PySide6/QtWidgets.pyd", "Qt Widgets Python binding"),
    ("_internal/PySide6/QtSvg.pyd", "Qt SVG Python binding"),
    ("_internal/PySide6/Qt6Core.dll", "Qt Core runtime"),
    ("_internal/PySide6/Qt6Gui.dll", "Qt GUI runtime"),
    ("_internal/PySide6/Qt6Widgets.dll", "Qt Widgets runtime"),
    ("_internal/PySide6/plugins/platforms/qwindows.dll", "Windows platform plugin"),
    ("_internal/assets/ui/home.svg", "Offline vector icon"),
    ("_internal/Excel自动汇总助手/qt_ui/check.svg", "Checkbox icon"),
    ("_internal/assets/licenses/LGPL-3.0-only.txt", "Qt license text"),
]

MIN_PYD_COUNT = 5    # 至少应包含若干 .pyd 扩展
MIN_DLL_COUNT = 5    # 至少应包含若干 .dll 运行库
MIN_EXE_SIZE = 1024 * 1024  # 引导 EXE 至少 1 MB（正常约 5 MB）

USAGE = """用法：
    python tools/check_dist.py <发行目录> [--qt]

发行目录必填，没有默认值。集中存放后的位置示例（在项目目录中执行）：
    python tools/check_dist.py ..\\本地产物\\Excel智能汇总助手\\dist\\<release>\\E助手

也可显式传入其他发行目录（例如迁移前的旧 dist 目录）。本脚本只读，不修改任何文件。

退出码：0 = 通过，1 = 缺失关键运行文件，2 = 用法错误。"""


def _fmt_size(n):
    return f"{n / (1024 * 1024):.2f} MB"


def check_integrity(dist_dir, *, qt=False):
    """检查发行目录完整性，返回 (ok: bool, report_lines: list)。"""
    lines = []
    ok = True

    lines.append("=" * 64)
    lines.append("E助手 v1.5.4 — 发行完整性检查")
    lines.append("=" * 64)
    lines.append(f"发行目录：{dist_dir}")
    lines.append("")

    if not os.path.isdir(dist_dir):
        lines.append(f"FAIL: 发行目录不存在：{dist_dir}")
        return False, lines

    # 1. 关键文件逐项检查
    lines.append("[关键文件]")
    for rel, desc in [*REQUIRED_FILES, *(QT_REQUIRED_FILES if qt else [])]:
        path = os.path.join(dist_dir, rel)
        if os.path.isfile(path) and os.path.getsize(path) > 0:
            lines.append(f"  PASS  {rel}  ({_fmt_size(os.path.getsize(path))})")
        else:
            lines.append(f"  FAIL  缺失：{rel}  ——  {desc}")
            ok = False

    if qt:
        # Windows Qt uses OS ICU; foreign ICU collected from PATH shadows it.
        for rel in ("_internal/icuuc.dll", "_internal/icudt78.dll"):
            if os.path.exists(os.path.join(dist_dir, rel)):
                lines.append(f"  FAIL  不应随包携带外部 ICU：{rel}")
                ok = False

    # 2. 可选文件（仅告警）
    lines.append("")
    lines.append("[次要运行库]")
    for rel, desc in OPTIONAL_FILES:
        path = os.path.join(dist_dir, rel)
        if os.path.isfile(path) and os.path.getsize(path) > 0:
            lines.append(f"  PASS  {rel}")
        else:
            lines.append(f"  WARN  缺失：{rel}  ——  {desc}")

    # 3. 引导 EXE 大小
    exe_path = os.path.join(dist_dir, "E助手.exe")
    if os.path.isfile(exe_path):
        exe_size = os.path.getsize(exe_path)
        if exe_size < MIN_EXE_SIZE:
            lines.append(f"")
            lines.append(f"FAIL: EXE 体积异常（{_fmt_size(exe_size)} < 1 MB），可能为损坏产物")
            ok = False

    # 4. 扩展与运行库数量统计
    internal = os.path.join(dist_dir, "_internal")
    pyds = [f for f in os.listdir(internal) if f.endswith(".pyd")] if os.path.isdir(internal) else []
    dlls = [f for f in os.listdir(internal) if f.endswith(".dll")] if os.path.isdir(internal) else []
    lines.append("")
    lines.append("[数量统计]")
    lines.append(f"  _internal/.pyd 数量：{len(pyds)}（最低 {MIN_PYD_COUNT}）")
    lines.append(f"  _internal/.dll 数量：{len(dlls)}（最低 {MIN_DLL_COUNT}）")
    if len(pyds) < MIN_PYD_COUNT:
        lines.append(f"FAIL: .pyd 数量不足 {MIN_PYD_COUNT}，发行物不完整")
        ok = False
    if len(dlls) < MIN_DLL_COUNT:
        lines.append(f"FAIL: .dll 数量不足 {MIN_DLL_COUNT}，发行物不完整")
        ok = False

    lines.append("")
    lines.append("=" * 64)
    lines.append("结论：" + ("PASS — 发行目录完整" if ok else "FAIL — 发行目录不完整"))
    lines.append("=" * 64)

    return ok, lines


def _print_usage(problem=None):
    print(USAGE)
    if problem:
        print(f"错误：{problem}")


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]

    if not argv:
        _print_usage("缺少发行目录参数。")
        return 2

    if argv[0] in ("-h", "--help"):
        _print_usage()
        return 0

    qt = len(argv) == 2 and argv[1] == "--qt"
    if argv[0].startswith("-") or (len(argv) > 1 and not qt):
        _print_usage(f"参数不合法：{' '.join(argv)}")
        return 2

    dist_dir = os.path.abspath(argv[0])
    ok, lines = check_integrity(dist_dir, qt=qt)
    for line in lines:
        print(line)

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
