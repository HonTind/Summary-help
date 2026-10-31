# -*- mode: python ; coding: utf-8 -*-

import os
import sys


def _find_runtime_dll(name):
    """从当前构建 Python 环境定位 MSVC 运行库，返回绝对路径；找不到返回 None。"""
    names = [name]
    if name.lower() != name:
        names.append(name.lower())
    if name.upper() != name:
        names.append(name.upper())

    roots = []
    for attr in ("base_prefix", "prefix"):
        root = getattr(sys, attr, None)
        if root and root not in roots:
            roots.append(root)
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    if exe_dir not in roots:
        roots.append(exe_dir)

    seen = set()
    for root in roots:
        for candidate in names:
            path = os.path.join(root, candidate)
            key = os.path.normcase(os.path.abspath(path))
            if key in seen:
                continue
            seen.add(key)
            if os.path.isfile(path):
                return os.path.abspath(path)
    return None


_VCRUNTIME140_1 = _find_runtime_dll("VCRUNTIME140_1.dll")
if _VCRUNTIME140_1 is None:
    sys.stderr.write(
        "构建中止：未能在当前 Python 环境（sys.base_prefix/sys.prefix/"
        "可执行文件目录）找到 VCRUNTIME140_1.dll。\n"
        "请使用包含 MSVC 运行库的 Python 3.12 环境构建，"
        "避免产出缺少该运行库的不完整软件。\n"
    )
    raise SystemExit(1)


a = Analysis(
    ['start_gui.py'],
    pathex=[],
    binaries=[(_VCRUNTIME140_1, ".")],
    datas=[(os.path.join(SPECPATH, 'assets', 'app-icon.ico'), 'assets'),
           (os.path.join(SPECPATH, 'assets', 'app-icon.png'), 'assets'),
           (os.path.join(SPECPATH, 'assets', 'ui'), 'assets/ui'),
           (os.path.join(SPECPATH, 'assets', 'licenses'), 'assets/licenses'),
           (os.path.join(SPECPATH, 'e_help', 'qt_ui', 'check.svg'), 'e_help/qt_ui')],
    hiddenimports=['openpyxl', 'openpyxl.styles', 'openpyxl.utils', 'openpyxl.cell', 'xlrd', 'xlrd.book', 'xlrd.sheet', 'xlrd.compdoc'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtNetwork', 'PySide6.QtQml', 'PySide6.QtQuick'],
    noarchive=False,
    optimize=0,
)
# Qt uses the unversioned Windows system ICU ABI, not Poppler/Conda ICU
# accidentally discovered on the development PATH. Leave OS ICU to Windows.
a.binaries = [entry for entry in a.binaries
              if os.path.basename(entry[0]).lower() not in {'icuuc.dll', 'icudt78.dll'}]
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='E-help',
    icon=os.path.join(SPECPATH, 'assets', 'app-icon.ico'),
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
