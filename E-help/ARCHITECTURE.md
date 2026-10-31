# E-help架构与维护说明

## 运行边界

本项目是完全离线的 Python/PySide6 桌面程序。`start_gui.py` 默认调用 `qt_ui.main()`；核心包名和层级保持 `e_help/`，开发 Agent 的模型接口不进入产品。明确使用 `--legacy-ui` 时才加载旧 `gui.main()` / `desktop_ui.UnifiedApp`。

2026-10-14用户确认迁移并要求继续完成。正式链是 `start_gui.py → qt_ui.main → ApplicationWindow → Backend → 原业务模块`。`application.py` 管理 Qt 控件与任务生命周期，`backend.py` 适配原引擎/数据转换/独占输出且无 Qt/Tk 依赖，`dialogs.py` 承接字段、组合键和文件映射。原业务算法不变。Qt Widgets Essentials 依赖进入正式 requirements，两份 spec 携带本地图标与许可，不引入 WebEngine 或服务。

`preview_qt_ui.py → PreviewWindow → DemoBackend` 仍是纯内存视觉样板；`preview_qt_production.py` 使用正式链但注入临时路径与合成输入。以下 Tk 链图保留作为兼容入口说明。

```text
start_gui.py → gui.main() → desktop_ui.UnifiedApp
                            ├─ 常规汇总适配 → gui.SummaryApp / _WorkerThread
                            └─ 月度/对账适配 → workbench_gui.WorkbenchController
                                                 ├─ job_runner（线程与事件）
                                                 ├─ monthly / reconciliation（计算）
                                                 ├─ workbench_view（预览与记录计数）
                                                 └─ workbench_export（生成 Excel）
```

## 模块职责与扩展位置

| 模块 | 单一职责与约束 |
| --- | --- |
| `desktop_ui.py` | 导航与嵌入式布局、任务互斥、帮助与反馈、统一记录页。功能页面切换不启动后台任务。 |
| `gui.py` | 兼容常规汇总窗口、字段配置/映射对话框、旧任务控制及正式入口。保留既有调用入口。 |
| `gui_style.py` | 原常规界面的颜色、字体、卡片、按钮等视觉组件。字体缓存属于 Tk 解释器，不访问业务数据。 |
| `workbench_gui.py` | 月度/对账输入控件、UI 状态、模板操作和任务调度；只在 Tk 线程更新控件。 |
| `job_runner.py` | 单个工作台的后台线程、取消和事件队列。无 Tk 依赖；不会自行写历史记录。 |
| `workbench_view.py` | 将计算结果转成预览行、说明文字及最小历史计数。无 Tk、文件写入或业务计算。 |
| `analysis.py` | 工作台输入规则、Excel 读取、金额与日期校验、来源追溯。读取双份 xlsx 以核对公式缓存；保留日期原序列处理。 |
| `monthly.py` | 月度结果模型与分组聚合。月份作为标签，只有明确选择日期字段才筛选。 |
| `template_store.py` | 月度模板校验、加载与原子写入；不保存旧账单路径或内容，不迁移旧格式。 |
| `reconciliation.py` | 客户总额与单据逐笔两种对账。重复单据保留歧义，不能擅自取第一项。 |
| `workbench_export.py` | 可追溯工作台报表、精确文本金额、公式注入防护、独占输出文件名及取消清理。 |
| `excel_reader.py` / `processor.py` / `excel_writer.py` | 原常规汇总的读取、排序和导出，保持动态字段与已有计算口径。读取完成、跳过和异常时都释放工作簿。 |
| `field_mapper.py` | 原常规字段识别及配置；仍维护现有进程内共享配置对象。 |
| `history.py` | 四类任务的最小本地记录、兼容旧记录、500 条保留上限与进程内锁。 |
| `utils.py` / `config.py` | 既有数据路径与产品常量；资源路径和可写数据路径不能混用。 |
| `app_icon.py` / `assets/` | 随程序分发的只读图标，从模块位置定位，支持源码与 PyInstaller。 |
| `offline_help.py` | 内置教程与静态公开联系方式，不接入网页/API。 |

## 为什么保留两条计算路径

常规汇总使用自动识别、动态字段和历史浮点计算/数值单元格输出；月度/对账使用明确规则、有序组合键、Decimal 精确计算与精确文本单元格。两者并非同一种算法的简单重复。将它们强行合并可能改变金额、字段优先级和输出格式，因此本轮保持独立。

`gui.py` 仍包含较多旧对话框。它们依赖原字段全局配置、Tk 状态及外部调用入口。样式已独立，后续增加新页面应通过独立视图与控制器接入，不往旧大类继续堆功能；进一步拆对话框需覆盖相应的交互、配置缓存和错误提示契约。

## Qt迁移接口与分阶段边界

第一阶段Qt模块：

- `qt_ui/demo.py`：`PreviewRequest/TableData/PreviewData`和真实已实现的内存演示后端。它不是外部模型API，也不是正式汇总算法；`TableData.footer`仅展示，不属于数据行。
- `qt_ui/assets.py`：本地许可SVG/软件logo，多DPR图标，不创建应用、不访问网络。
- `qt_ui/theme.py`：应用内浅色调色板、逻辑尺寸、字体、QSS；不改变操作系统设置。
- `qt_ui/widgets.py`：表格/分页/折叠区；`qt_ui/window.py`：原生导航与演示交互，UI状态与模型分离。
- `tools/test_qt_ui.py`：独立无窗口Qt验收，不强迫现有Tk测试安装Qt。

Qt正式后端已经直接接入以下无GUI模块，不继承Tk `WorkbenchController`：

1. 月度/对账：`analysis.SheetRules`、`monthly.make_template/run_monthly`、`reconciliation.run_reconciliation`。`#2`是UI文本约定，转换为1起始整数列号；组合键顺序与最多8项约束保留。月份仅在明确设置date_column后筛选。
2. 结果：`workbench_view.build_preview/history_counts`提供现有显示/记录口径；`workbench_export.export_result`提供只读输入、独占新文件名、精确文本、防公式注入及不完整确认边界。Qt只做显示，不把展示合计混进导出数据。
3. 生命周期：复用`job_runner.JobRunner`，Qt主线程定时消费progress/done/cancelled/error；工作线程不调用控件。导出使用`accept_completed=True`，不能把保存返回后晚到取消误报失败。线程停止后二次drain，终态历史只记一次；运行中冻结输入，安全停止后关闭。
4. 配置/历史：直接复用`template_store`和`history`的校验、原子写、坏文件保护、500条上限与白名单字段。用户明确确认同名覆盖/删记录；历史保存失败不改变已完成导出状态。不保存整份result.config、异常原文或输入名单。
5. 常规汇总另做适配：`excel_reader.aggregate_with_report → processor.sort_by_amount → excel_writer.generate_summary`，skipped不空须阻止导出；保留字段共享配置与原数值格式。该路径没有合作取消，不能套用月度的取消承诺；不直接复用Tk模块的daemon工作线程。
6. 路径：`utils`以项目/EXE目录定位可写数据；资源路径独立。get_config_dir/get_history_dir即使读取也可能创建目录，第一阶段绝不调用。真实迁移不移动用户数据或授权文件。

字段映射、规则、记录和 About 已用 Qt 视图承接；正式入口与 spec 已适配。新构建使用独立 release 验证，旧 EXE/快捷方式保留。不要启动两个正式实例写同一数据文件；历史锁目前只覆盖同进程。Git 提交和发布仍由用户批准。

`assets/licenses/` 保存上游许可及源码渠道，两份 spec 随包携带。onedir 使用独立动态库。对外分发前仍核对实际组件及对应源码版本；本地验收不是对外分发合规保证。[Qt for Python许可说明](https://doc.qt.io/qtforpython-6/licenses.html)、[Qt高DPI说明](https://doc.qt.io/qt-6/highdpi.html)。本轮本地构建不代表发布 Release。

## 兼容规则

Qt Windows 运行库依赖操作系统 ICU 的未版本化导出。2026-10-14 的隔离启动揭示 PyInstaller 从开发 PATH 的 Poppler 收集了不同 ABI 的 `icuuc.dll`/`icudt78.dll`，导致 `QtWidgets` DLL 导入失败；两份 spec 在分析后过滤这两项，仅由 Windows 提供系统 ICU，不复制/替换系统 DLL。`tools/test_qt_production.py` 对两份 spec 的过滤及 `check_dist.py --qt` 污染拒绝设有回归。最终唯一 `qt-offline-20261014-03` 构建通过三种路径启动；负例错误弹窗的自动确认未完成，保留工具 FAIL，不把超时写成 PASS。

- `gui` 继续导出原样式函数与常量；两个根窗口缓存属性从 `gui_style` 动态读取，避免产生过期副本。新视觉代码可直接引用 `gui_style`。
- `workbench_gui.JobRunner` 保留为同一类的别名。预览状态和取消语义保持原样：已保存完成的导出不能被之后的取消误报为失败。
- `monthly` 继续导出模板函数；新模板存储代码直接使用 `template_store`。保留 `monthly.os` 供既有原子保存故障注入测试使用。
- 历史与模板分别保持各自 JSON 协议；坏文件不自动清空，未知版本不覆盖。不要为了共用一个 JSON 工具抹平两者不同的异常策略。
- `field_mapper` 配置列表/字典的原地更新是现存依赖，不能简单替换成新对象。统一记录仅保证进程内互斥，暂不支持多实例并发写同一文件。
- 两份 spec 的少量公共构建代码暂时保留，修改运行库或资源配置时应同时核验；没有为此引入额外构建框架。

## 验证与开发工具

常规回归入口（使用临时合成数据）：

```powershell
python -B -m unittest discover -s e_help/tests -p 'test_*.py' -q
```

`test_architecture_boundaries.py` 验证无 Tk 的结果转换、精确金额显示、缺失/零金额、重复歧义及无副作用导入；`test_legacy_workbook_resources.py` 覆盖成功、跳过和异常时的资源释放；`test_launch_isolation.py` 使用模拟进程验证开发工具隔离，绝不启动真实 EXE。

`tools/test_launch.py <onedir发行目录>` 是需要实际启动程序时单独运行的工具：仅从原发行目录复制 EXE 和 `_internal`，不读取或迁移旁边的用户数据。测试副本位于 `本地产物/E-help/build/launch-check-*`；只终止本次持有的进程，不按名称关闭其他实例。副本和日志保留供排查，不递归清理历史产物。它只支持 onedir；缺失依赖的负面检查必须真实失败退出，超时不会被当作通过。

旧 `verify_*.py`、压力与稳定性脚本有历史路径假设，不纳入默认 discovery；需要时先读脚本，在隔离环境运行。不要对用户正在使用的发行目录运行破坏性诊断。

构建使用 `tools/build_release.ps1 -Release <唯一名称>`；静态导入的新模块随 PyInstaller 分析进入包，不改变原项目路径、入口或运行时依赖。旧 EXE 不随源码变化，验收源码与打包软件应分别说明版本来源。
