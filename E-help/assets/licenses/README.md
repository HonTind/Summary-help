# 第三方许可与本地构建说明

E-help界面使用 PySide6-Essentials / shiboken6 6.12.0（Qt for Python）和现有 openpyxl、xlrd。Qt Widgets、GUI、Core、SVG 库以独立动态库放在 onedir 的 `_internal` 中，不静态链接；本软件不限制为调试 Qt 修改而进行的逆向工程。

Qt for Python 的上游许可文本随本目录及库包一起保留。标准界面图标采用 Feather 4.29.2，其完整 MIT 许可见 `assets/ui/FEATHER-LICENSE.txt`。软件 logo 使用用户提供的素材。

本仓库中的 Python 源码、两份 spec、requirements.txt 与 tools/build_release.ps1 是可复现构建入口。安装 requirements.txt 后，用独立 Python 环境运行构建脚本；onedir 允许在符合兼容版本与依赖约束的情况下替换动态库。分享安装包时应随附本许可目录及上游库的源码取得渠道，不应只分享 EXE。Qt 源码及构建资料：

- https://code.qt.io/cgit/qt/qtbase.git/
- https://code.qt.io/cgit/qt/qtsvg.git/
- https://code.qt.io/cgit/pyside/pyside-setup.git/
- https://download.qt.io/official_releases/QtForPython/
- https://doc.qt.io/qtforpython-6/licenses.html

当前构建仅供本地验收，没有发布 Release。对外分发前按实际包内容核对所有第三方通知与对应源码版本；本说明不替代完整分发合规审查。优先使用 onedir，onefile 在分发时需要额外核对替换库和源码提供安排。
