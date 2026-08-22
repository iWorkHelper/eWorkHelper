# eWorkHelper 开发日志

本日志只保留重要功能、架构、兼容性、版本和发布准备变更。临时调试过程、已失效方案及逐次尝试不作为长期记录保存。

## 2026-08-22：v1.1.260822.6 正式发布准备

- 按统一版本规则为最终正式发布构建分配 `1.1.260822.6`；完整产品版本写入 `AssemblyInformationalVersion`，ClickOnce/VSTO 数值发布字段使用合法的 `1.1.0.6`。此前同日候选构建未发布，最终内容变化后递增 Revision，未复用既有完整版本号。
- 从公开项目文件移除本地临时签名证书文件名与指纹绑定，补充 Visual Studio 性能文件、二进制日志、环境文件和本地 JSON 配置忽略规则。
- 更新 README、Release Notes、功能说明与当前版本说明；Debug / Any CPU 与 Release / Any CPU Rebuild 均通过，Error 0、Warning 0，程序集 ProductVersion 均为 `1.1.260822.6`。
- 首轮产物审计发现旧本地证书 Subject 含环境身份信息，因此废弃该产物；最终构建改用通用 `CN=eWorkHelper Release` 临时自签名证书，构建后销毁临时私钥。候选资产复扫通过，PDB 因包含本机源码路径不进入 Release。
- 项目没有独立自动化测试；静态与 manifest 完整性检查通过，Excel F5/GUI 和真实工作簿场景仍按回归清单标记为待人工验证。

## 2026-08-22：新增取消合并并填充

- 在现有 `eWorkHelper` Ribbon“数据工具”Group 中新增第二个正式大按钮“取消合并并填充”，未增加 Tab、测试 Group 或临时入口；原“批量过滤”事件和实现未修改。
- 新增 `UnmergeAndFillService.cs`：读取当前 Selection 的全部 Areas，通过范围级 `MergeCells` 判断与递归分块剪枝发现 MergeArea；以工作簿/工作表外部绝对地址去重，每个 MergeArea 只处理一次。
- 每个区域在 `UnMerge()` 前读取左上角内容；普通内容用 `Value2` 整块填充，空值不写入字符串，公式用 `FormulaR1C1` 整块填充并保留 Excel 相对引用复制语义。
- 临时关闭 `ScreenUpdating`、`EnableEvents`，通过嵌套 `finally` 确保两项状态均尝试恢复；Calculation 不修改。异常由 Ribbon 使用现有 MessageBox 规范反馈，项目未引入新的日志或架构框架。
- 同步功能设计、开发入口、VSTO 架构、回归清单、Release Notes 和版本元数据。该功能构成新的功能阶段，Minor 更新为 `1.1`，同日 Revision 递增，产品版本为 `1.1.260822.4`。
- Debug / Any CPU 与 Release / Any CPU Rebuild 均通过，Error 0、Warning 0；两个程序集 ProductVersion 均为 `1.1.260822.4`。
- 静态检查确认 Ribbon 无 Test/Demo/Debug 符号，核心仅有一个原生 `UnMerge()` 调用点，公式与状态恢复路径存在。仓库没有自动化测试项目；Excel F5/GUI、真实合并区域行为、公式相对引用、格式和受保护工作表异常路径仍需按回归清单人工验证。
- 已知限制：Excel/VSTO 不为该多区域操作提供可组合的事务撤销单元；若后续区域处理失败，已完成区域无法自动回滚，但 Application 临时状态会恢复并向用户报告错误。

## 2026-08-22：批量过滤窗口自动加载当前筛选内容

- 修改 `BatchFilterService.cs`、`BatchFilterForm.cs`、开发说明、功能说明、回归清单、Release Notes 和版本元数据。
- 仅当目标字段 `Filter.On` 为 true 时加载，避免有筛选箭头但字段未实际过滤时误填整列。
- 运行时状态保存 Workbook、Worksheet、AutoFilter Range 边界、Field、MatchMode、原始 Conditions 和 Filter 签名；上下文和签名均一致时优先恢复 eWorkHelper 原始规则。
- 外部 Excel 筛选通过目标 DataRange 的 `SpecialCells(xlCellTypeVisible)` 获取可见 Areas，每个 Area 批量读取 `Value2`，在内存中去空、按首次出现顺序进行不区分大小写去重，并以“等于”模式回填。
- 不调用 `ShowAllData`，不取消或修改其他字段 Filter，不使用 Hidden、辅助列或临时 Sheet；零可见行保持空文本并显示状态。
- 本次为现有功能扩展，Major/Minor 保持 `1.0`，同日 Revision 更新，产品版本为 `1.0.260822.3`。
- 编译、自动验证和 Excel F5 实测结果在完成后补充。

## 2026-08-22：新增 Ribbon 关于入口

- 在 `eWorkHelper` Tab 最末端新增独立“关于”Group 和大按钮，不改变“数据工具”及批量过滤行为。
- 点击后显示简短插件介绍和完整产品版本；版本从当前程序集的 `AssemblyInformationalVersion` 动态读取，不维护 UI 硬编码版本。
- 本次属于小范围功能补充，不进入新的功能阶段，Major/Minor 保持 `1.0`；同日 Revision 递增，产品版本更新为 `1.0.260822.2`。
- 同步 README、Release Notes、开发入口、架构说明和 Ribbon 回归清单。
- Debug / Any CPU 与 Release / Any CPU Rebuild 均通过，Error 0、Warning 0；两个生成程序集的 ProductVersion 均为 `1.0.260822.2`。
- Ribbon 静态验证确认 Group 顺序为“数据工具 → 关于”，关于事件绑定、动态 `AssemblyInformationalVersion` 读取和原批量过滤事件均正常，UI 不包含硬编码版本号。
- 本次未执行 Excel F5/GUI 人工验证；需要人工确认 Ribbon 实际布局和关于窗口显示效果。

## 2026-08-22：Project cleanup / documentation consolidation / privacy sanitization

### 目标与范围

- 清除生成物、本地痕迹和已失效测试命名，不改变批量过滤业务行为。
- 将碎片化的架构、功能、测试和发布说明整理为按职责维护的文档体系。
- 收敛开发日志，移除已被最终实现替代的旧方案和调试流水账。
- 重新执行全仓库隐私扫描、Debug/Release 构建及 VSTO/Ribbon/资源静态验证。

### 主要整理

- 将承担正式功能的 `RibbonTest.*` 重命名为 `MainRibbon.*`，并将类名 `MaimRibbon` 修正为 `MainRibbon`；控件、事件和过滤逻辑不变。
- 删除 ThisAddIn 与 AssemblyInfo 中未使用的 using，不修改生命周期或程序集行为。
- 新增 `Architecture/VstoArchitecture.md`、`Features/BatchFiltering.md`、`Testing/RegressionChecklist.md`、`Release/ReleaseChecklist.md`。
- `Development.md` 收敛为开发流程、文档导航和当前基线入口；一次性安全审计结论合并到长期发布检查清单。
- `.gitignore` 补充 dump、诊断跟踪、IDE 插件缓存、临时导出与 Publish Profile 规则。

### 安全与验证

- 初始扫描发现的本机绝对路径仅位于 ignored 的构建中间文件中，通过删除生成目录处理；源文件和现有 Git 历史未发现真实凭据、证书或身份信息。
- 最终候选文件扫描未发现真实身份、组织、客户、本地路径、Secret、私钥、证书或内部网络信息；程序集公开版本号等格式误报已人工分类确认。
- Debug / Any CPU Rebuild 通过，Error 0、Warning 0，ProductVersion 为 `1.0.260822.1`。
- Release / Any CPU Rebuild 通过，Error 0、Warning 0，ProductVersion 为 `1.0.260822.1`。
- MainRibbon partial/Designer/Resource、批量过滤事件、Ribbon Tab、ThisAddIn Startup/Shutdown 和项目引用静态验证全部通过，旧测试符号为零。
- 当前仓库没有独立自动化测试项目，因此无自动化测试可运行；Excel F5/GUI 本次未执行，长期人工回归项见 `Testing/RegressionChecklist.md`。

## 2026-08-22：建立统一版本并更新当前程序版本

- 建立唯一格式 `Major.Minor.Date.Revision`，完整规则见 [`Versioning.md`](Versioning.md)。
- 当前产品版本确定为 `1.0.260822.1`：Major/Minor 保持 `1.0`，Date 为 `260822`，当日 Revision 为 `1`。
- 完整产品版本写入 `AssemblyInformationalVersion`。由于 .NET Framework 数值版本字段单段最大为 `65535`，`AssemblyVersion` 与 `AssemblyFileVersion` 保持兼容值 `1.0.0.0`，不用于对外展示。
- 版本调整后 Debug/Release Rebuild 均通过，Error 0、Warning 0；两个程序集的 ProductVersion 均为 `1.0.260822.1`。

## 2026-08-22：首次公开发布准备

- 创建适用于 Visual Studio、C#、VSTO、.NET Framework 和 ClickOnce 的 `.gitignore`。
- 创建公开 README、MIT License、Release Notes 和安全审计记录。
- 将程序集身份信息替换为中性的 `eWorkHelper Contributors`，移除项目文件中的开发证书指纹，公开源码默认不绑定签名凭据。
- 初始目录没有 Git 历史；完成安全扫描后初始化本地仓库并创建首次公开发布提交。仓库当时未配置 remote，因此未 Push、Tag 或创建 GitHub Release。
- Debug/Release VSTO Rebuild 均通过，Error 0、Warning 0；签名只通过本地 MSBuild 进程临时提供，未进入仓库。

## 2026-08-22：批量文本过滤正式实现

- 在 `eWorkHelper` Ribbon 的“数据工具”Group 中提供正式“批量过滤”功能，移除早期环境验证用测试 Group、按钮、事件和提示。
- 支持“等于”“不等于”“包含”“不包含”四种不区分大小写的匹配方式，默认“包含”；输入按行 Trim、忽略空行并精确去重。
- 最终实现采用 C# 内存匹配和 Excel 原生 AutoFilter，不写改单元格、不创建辅助列/工作表、不直接隐藏行。
- 支持 ListObject、已有 AutoFilter 和用户选择标题行建立的普通区域；只操作目标 Field，保留其他字段筛选条件。
- 无匹配使用唯一不存在值产生原生零行结果；全部匹配只清除目标字段条件；“清除本次过滤”限于当前 Add-in 会话最近处理的同一目标。
- Debug 构建与条件匹配静态/样例验证通过。长期人工回归范围见 [`Testing/RegressionChecklist.md`](Testing/RegressionChecklist.md)。

## 2026-08-21 至 2026-08-22：VSTO 基线与开发约束

- 确认项目为 Excel VSTO Add-in，目标框架 .NET Framework 4.8，Visual Studio 2022，Excel Host，LoadBehavior 为 `3`。
- 建立唯一 Ribbon Designer 和基础事件链路；用户曾人工确认 F5 启动 Excel、Add-in 加载、Ribbon 显示及按钮事件可用。
- 建立 [`DevelopmentConstraints.md`](DevelopmentConstraints.md)，固定先分析、最小修改、完成验证、代码与文档一致等长期约束。
- 公开仓库不保存本地开发证书信息；换机或换用户时应在受保护环境重新提供本地签名配置。
