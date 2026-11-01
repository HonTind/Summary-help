# Summary-help · E-help

项目源码位于 `E-help/`，Python 包名为 `e_help`，打包程序名称为 `E-help.exe`。
新安装使用 `E-help-data/` 保存设置与历史记录；已有旧版数据目录时继续读取原目录，不自动移动或删除数据。

本地运行的 Excel 汇总与账单对账工具，使用 Python 与 PySide6 构建 Windows 桌面界面。业务文件在本机处理；在线问题反馈仅在用户主动确认后通过系统浏览器打开，不自动上传账单。

## 功能

- 常规汇总：识别表头、配置字段映射、预览并导出结果。
- 月度汇总：复用规则模板，支持组合键与可选日期筛选。
- 账单对账：按客户总额或单据逐笔比较，查看差异与来源明细。
- 本地任务记录、规则与模板、内置使用教程。

处理结果请结合来源明细复核；建议保留原始文件备份。具体计算口径见 [功能说明](E-help/FEATURES.md)。

## 从源码运行

目标环境为 Windows 10/11；当前回归验证使用 Python 3.12。首次安装依赖需要联网，日常业务处理不需要在线服务。

```powershell
git clone https://github.com/HonTind/Summary-help.git
cd Summary-help/E-help
python -m venv .venv/qt-ui
.venv/qt-ui/Scripts/python.exe -m pip install -r requirements.txt
.venv/qt-ui/Scripts/python.exe -B start_gui.py
```

请保持 `E-help/` 项目目录与 `e_help/` Python 包名不变，避免破坏路径引用。GitHub 的源码 ZIP 不是可直接运行的 EXE；没有 Git 时，也可下载源码 ZIP、完整解压后，在 `E-help/` 目录执行上面的虚拟环境、安装依赖和启动命令。

## 使用已打包的软件

如果你已经取得 Windows 64 位免安装软件包，请完整解压，再打开其中的 `E-help/E-help.exe`。无需另外安装 Python 或 Microsoft Excel。EXE 与同级 `_internal` 文件夹必须一起保留，不能只复制 EXE 或直接在压缩软件中运行。

若下载文件附带 `双击这里启动E-help.cmd` 和 `启动工具-请勿单独删除.ps1`，将它们与对应版本的软件 ZIP 放在同一文件夹，再双击 CMD 入口。入口会校验软件包并在首次运行时解压。脚本受系统限制时，可使用上述手动解压方式。

## 首次使用

建议先用少量样例文件检查规则，核对结果后再处理正式账单。

1. **常规汇总**：在工作台选择常规汇总，选择账单文件夹，核对字段映射和处理规则，预览后导出。
2. **月度汇总**：选择文件、模板和月份，核对关键字段与金额字段。需要限定记录日期时，启用日期筛选；仅选择月份不代表已过滤数据。
3. **账单对账**：分别选择左右账单与规则，选择客户总额或单据逐笔模式，核对匹配键、金额字段和容差，再预览差异并导出。

导出前检查金额合计、来源行数和异常清单，保留原始账单。任务记录可查看历史处理情况；规则与模板用于保存可复用配置，界面中的使用教程提供进一步指引。

设置和历史默认保存在程序旁的 `E-help-data/`，已有旧版数据目录时兼容读取。升级或清理前，请备份需要保留的数据和导出文件。当前版本不要求机器码激活，也不限制汇总次数。

## 常见问题与反馈

- **无法启动**：检查是否完整解压、`_internal` 是否存在，避免混入其他版本的 DLL。
- **无法保存**：将整个软件文件夹和输出目录放到当前账户可写的位置。
- **结果不符**：先缩小到少量文件，检查表头、字段映射、日期筛选、金额格式和对账规则，并与原始账单复核。
- **反馈问题**：通过界面确认后打开[腾讯问卷](https://wj.qq.com/s2/28242546/gc6x/)。反馈网页需要联网，不影响本地业务功能；请使用脱敏文字和虚拟样例。

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

产物位于仓库根目录 `本地产物/E-help/{build,dist}/`，不纳入 Git。onedir 版本需要保留 EXE 与同级 `_internal`，不能只复制 EXE。正式分发前应验收构建结果并核对第三方许可义务。

## 目录与文档

- [项目详细说明](E-help/README.md)
- [架构说明](E-help/ARCHITECTURE.md)
- [贡献与问题反馈](CONTRIBUTING.md)
- [迁移来源](MIGRATION.md)
- [GPL-3.0-only](LICENSE) 与 [许可适用范围](LICENSING.md)

## 隐私与许可证

请勿在 Issues、Pull Requests 或问卷提交真实账单、客户信息、密码、API 密钥、私钥或个人路径；使用脱敏截图和虚拟样例。

项目原创代码按 GPL-3.0-only 提供，第三方依赖与 Feather 图标保留各自许可。品牌 Logo 不包含在本次 GPL 授权范围内，具体边界及来源核实事项见 [LICENSING.md](LICENSING.md)。本项目不以开源声明替代第三方素材的权属确认。
