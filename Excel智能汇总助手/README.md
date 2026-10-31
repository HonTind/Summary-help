# E助手项目说明

本目录包含应用入口、业务模块、资源、测试和构建配置。安装与运行命令见 [仓库首页](../README.md)。

## 业务功能

常规汇总、月度汇总、客户总额对账与单据逐笔对账共用主窗口。支持组合键、日期筛选、左右独立规则、容差、字段映射、模板、结果预览及导出。具体计算与异常处理口径见 [FEATURES.md](FEATURES.md)。

任务记录保存在本机，包含任务类型、状态、统计数量和输出路径，不记录账单正文。输出路径可能包含个人信息，请勿公开运行目录。删除记录不会删除对应 Excel 文件。

## 入口

- 正式界面：`python -B start_gui.py`。
- 旧 Tk 兼容界面：`python -B start_gui.py --legacy-ui`。
- 隔离合成数据预览：`python -B tools/preview_qt_production.py`。
- 内存界面样板：`python -B tools/preview_qt_ui.py`，不处理真实账单，导出仅展示操作说明。

正式入口不会自动载入样例。隔离预览数据只写入仓库根目录的本地产物测试区域。源码更新不会自动替换旧 EXE。

## 测试与构建

从此目录运行 `python -B tools/test_free_suite.py`，命令名称保留历史兼容。核心和 Qt 测试使用合成数据，测试依赖和命令见首页。旧压力测试与手工验证脚本不属于默认测试套件。

构建统一使用 `tools/build_release.ps1`。保持当前目录结构不变，构建输出集中到仓库根目录 `本地产物/Excel智能汇总助手/`。脚本拒绝覆盖相同 release 名称。

完整 onedir 目录包含 EXE 与 `_internal`。分发前可运行 `python -B tools/check_dist.py <发行目录> --qt` 检查资源；不要将本机其他软件的同名 Qt/ICU DLL 混入发行包。

## 反馈与隐私

业务数据在本机处理。问题反馈按钮经用户确认后打开腾讯问卷，需要联网；不会自动附带账单或日志。取消反馈不影响汇总、对账或导出。

不要公开真实账单、客户资料、密码或个人文件路径。隐私检查工具 `tools/audit_repository.py` 仅做启发式扫描，不能证明不存在任何敏感信息。

## 架构与许可

模块边界见 [ARCHITECTURE.md](ARCHITECTURE.md)。不要随意重命名 `Excel自动汇总助手/`，避免破坏导入和数据路径依赖。

源码许可见 [LICENSE](../LICENSE) 与 [LICENSING.md](../LICENSING.md)。第三方库、Feather 图标及品牌素材的授权边界分别保留，不因公开迁移而改变。
