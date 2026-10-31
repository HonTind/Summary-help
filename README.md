# Summary-help · E助手

本地运行的 Excel 汇总与账单对账工具，使用 Python 与 PySide6 构建 Windows 桌面界面。业务文件在本机处理；在线问题反馈仅在用户主动确认后通过系统浏览器打开，不自动上传账单。

## 功能

- 常规汇总：识别表头、配置字段映射、预览并导出结果。
- 月度汇总：复用规则模板，支持组合键与可选日期筛选。
- 账单对账：按客户总额或单据逐笔比较，查看差异与来源明细。
- 本地任务记录、规则与模板、内置使用教程。

处理结果请结合来源明细复核；建议保留原始文件备份。具体计算口径见 [功能说明](Excel智能汇总助手/FEATURES.md)。

## 从源码运行

目标环境为 Windows 10/11；当前回归验证使用 Python 3.12。首次安装依赖需要联网，日常业务处理不需要在线服务。

```powershell
git clone https://github.com/HonTind/Summary-help.git
cd Summary-help/Excel智能汇总助手
python -m venv .venv/qt-ui
.venv/qt-ui/Scripts/python.exe -m pip install -r requirements.txt
.venv/qt-ui/Scripts/python.exe -B start_gui.py
```

请保持项目目录和中文 Python 包名不变，避免破坏已有路径引用。源码下载不是已打包的 EXE；本次迁移不包含安装包或 Release。

## 测试

在项目子目录内运行，测试仅使用合成数据：

```powershell
.venv/qt-ui/Scripts/python.exe -m pip install xlwt
$env:QT_QPA_PLATFORM = 'offscreen'
.venv/qt-ui/Scripts/python.exe -B tools/test_free_suite.py
```

测试入口名称为历史兼容名称。默认套件包括核心业务与两组 Qt 测试，不依赖签发工具或真实用户文件；旧压力脚本不自动运行。`xlwt` 仅用于测试生成旧格式样例，不是应用运行依赖。

## 打包

使用同一虚拟环境安装 PyInstaller，然后调用构建脚本；每次使用不同的 release 名称：

```powershell
.venv/qt-ui/Scripts/python.exe -m pip install pyinstaller
powershell -File tools/build_release.ps1 -Release local-001 -PythonExe .venv/qt-ui/Scripts/python.exe
```

产物位于仓库根目录 `本地产物/Excel智能汇总助手/{build,dist}/`，不纳入 Git。onedir 版本需要保留 EXE 与同级 `_internal`，不能只复制 EXE。正式分发前应验收构建结果并核对第三方许可义务。

## 目录与文档

- [项目详细说明](Excel智能汇总助手/README.md)
- [架构说明](Excel智能汇总助手/ARCHITECTURE.md)
- [贡献与问题反馈](CONTRIBUTING.md)
- [迁移来源](MIGRATION.md)
- [GPL-3.0-only](LICENSE) 与 [许可适用范围](LICENSING.md)

## 隐私与许可证

请勿在 Issues、Pull Requests 或问卷提交真实账单、客户信息、密码、API 密钥、私钥或个人路径；使用脱敏截图和虚拟样例。

项目原创代码按 GPL-3.0-only 提供，第三方依赖与 Feather 图标保留各自许可。品牌 Logo 不包含在本次 GPL 授权范围内，具体边界及来源核实事项见 [LICENSING.md](LICENSING.md)。本项目不以开源声明替代第三方素材的权属确认。
