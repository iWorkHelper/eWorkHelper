# eWorkHelper 开发说明

> 后续所有开发必须先阅读并遵守 [`DevelopmentConstraints.md`](DevelopmentConstraints.md)。

## 公开发布状态

- 首次公开源码版本准备为 `v1.0.0`，对应程序集版本 `1.0.0.0`。
- 根目录已提供 README、MIT License、Release Notes 和适用于 Visual Studio/VSTO 的 `.gitignore`。
- 公开项目不保存证书指纹、证书、私钥或用户级发布配置；完整 VSTO 清单构建需要维护者在受保护的本地/发布环境中提供临时签名配置。
- 发布前必须按 [`SecurityReview.md`](SecurityReview.md) 执行候选文件、Git 历史、差异和生成内容复查；任何安全检查未通过时禁止 Commit、Push 和 Release。
- 当前仓库不提供签名安装包，源码 Release 不应附带 `bin`、`obj`、DLL、PDB 或 Debug 输出。

## 已验证开发基线

以下能力已由用户人工验证通过：

- Visual Studio F5 可正常启动 Excel。
- eWorkHelper VSTO Add-in 可正常加载。
- 自定义 Ribbon 可正常显示。
- Ribbon 按钮事件链路可正常触发（历史验证事实；早期测试按钮现已移除）。

后续修改如影响上述链路，交付前必须重新验证。

## 当前项目基本结构

- `eWorkhelper.sln`：Visual Studio 2022 解决方案。
- `eWorkhelper.csproj`：Excel VSTO Add-in 项目，输出类型为类库。
- `ThisAddIn.cs` / `ThisAddIn.Designer.cs` / `ThisAddIn.Designer.xml`：VSTO Add-in 主机项及生成代码。
- `RibbonTest.cs` / `RibbonTest.Designer.cs` / `RibbonTest.resx`：本阶段的 Ribbon Designer 实现。
- `Properties`：程序集信息、资源和设置。

项目根命名空间及现有代码命名空间均为 `eWorkhelper`。

## 当前开发环境

- C#、.NET Framework 4.8。
- Visual Studio 2022，VSTO 4.0 项目系统。
- 宿主应用为 Windows 桌面版 Microsoft Excel。
- Office/Excel Interop 15.0 引用采用嵌入互操作类型。
- Debug 配置使用 VSTO 原生启动：从 Office 16.0 注册表安装路径定位 `excel.exe`，启动参数为 `/x`。
- 公开源码默认不绑定任何开发者证书，项目文件不保存证书指纹、私钥或签名凭据。

## 当前功能

当前 Ribbon 保留一个 `eWorkHelper` Tab，其中“数据工具”Group 提供正式的“批量过滤”功能。早期用于验证 VSTO 环境的“测试”Group、按钮和提示框已删除；VSTO、Excel 启动、Ribbon 加载及按钮事件链路已经人工验证通过的历史基线保持有效。

## 涉及文件

- `eWorkhelper.csproj`
- `RibbonTest.cs`
- `RibbonTest.Designer.cs`
- `RibbonTest.resx`
- `docs/Development.md`
- `docs/DevelopmentLog.md`

## 实现方案

采用 VSTO Ribbon Designer 单一实现方式。Excel Workbook Ribbon 中显示 `eWorkHelper` Tab，当前正式 UI 为“数据工具”Group 中的 `btnBatchFilter`“批量过滤”按钮。

`ThisAddIn_Startup` 和 `ThisAddIn_Shutdown` 保持默认空实现，不自行启动 Excel，不创建额外 COM Application。

公开仓库中的项目默认关闭 ClickOnce 清单签名，以避免绑定本地开发者凭据。如需制作签名安装包，维护者应在本机或受保护的发布环境中配置证书，且不得将证书、私钥、指纹或用户级发布配置提交到 Git。

## 调试方式

1. 在 Visual Studio 中打开 `eWorkhelper.sln`。
2. 选择 `Debug | Any CPU`。
3. 按 F5 或选择“开始调试”。
4. VSTO 应使用项目内的 Excel Host/Start Action 配置启动新的 Excel 实例。
5. 打开 `eWorkHelper` Tab，确认“数据工具”Group 和“批量过滤”按钮正常显示。
6. 在有效数据列中选择单元格后打开批量过滤窗口，验证原生 AutoFilter 流程。

## 验收标准

- 项目编译成功且不新增第三方依赖。
- F5 启动 Excel，并正常加载 eWorkHelper Add-in。
- Ribbon 显示一个 `eWorkHelper` Tab；Tab 内有“数据工具”Group 和正式“批量过滤”按钮，不再显示早期测试控件。
- Ribbon Designer 源文件、Designer 文件、资源文件及项目项匹配。
- GUI 行为如未能在当前环境执行，则必须由开发者按上述步骤人工验证。

## 批量文本过滤功能设计

### 功能目标与 UI

在现有 `eWorkHelper` Tab 的“数据工具”Group 中提供 `btnBatchFilter`“批量过滤”按钮。按钮从 `Globals.ThisAddIn.Application` 获取当前 Excel 上下文，识别目标列和数据区域后打开 `BatchFilterForm`。

窗体只展示自动识别的目标列，不允许手工修改列号；保留多行条件输入框，提供“等于”“不等于”“包含”“不包含”四个互斥 RadioButton（默认“包含”）、“应用过滤”、“清除本次过滤”和“关闭”按钮，并在窗体内显示检查行数及匹配行数。条件按行拆分、去除每行首尾空格、忽略空行，并以精确文本比较去除完全重复项；除首尾空格外不改写条件内容。

### 数据范围与匹配规则

目标列固定为启动功能时的 `ActiveCell.Column`。若活动单元格位于 `ListObject`，直接复用 `ListObject.Range` 及其原生 AutoFilter；否则若工作表已有 AutoFilter，目标列在 `Worksheet.AutoFilter.Range` 内时复用该区域，目标列不在其中时停止且不改变已有筛选。不得重新建立或清除用户已有筛选区域。

普通区域尚未启用 AutoFilter 时，先提示用户选择标题行，再通过 `Application.InputBox(Type: 8)` 让用户在工作表中选择。允许选择标题行中的一个或多个单元格，但必须只跨一行；取消时不修改工作表。服务用所选单元格的 `CurrentRegion` 判断左右边界和底部，最终建立的 AutoFilter Range 第一行严格使用用户所选行，绝不包含标题上方的报表名称、日期或说明。最初选择的目标列必须位于最终范围中。

服务一次性读取目标数据列 Range 的 `Value2`，将空值、数字、日期及其他值统一转换为文本，并使用 `OrdinalIgnoreCase` 比较。“等于”保留与任一输入条件完整相等的值；“不等于”排除与任一输入条件完整相等的值；“包含”保留包含任一输入条件的值；“不包含”只保留不包含所有输入条件中的任意一项的值，即任一条件被包含时都排除。匹配后对完整单元格值去重，再通过 `Range.AutoFilter`、完整值数组和 `xlFilterValues` 交给 Excel 原生筛选。条件数量不设人为上限。

### 非破坏性策略与状态恢复

批量过滤正式采用“C# 内存条件计算 + Excel 原生 AutoFilter”。过滤不修改单元格、不增加辅助列或工作表、不使用临时标记，也不直接设置任何行的 Hidden 状态。原生筛选箭头会保留，用户可以继续使用 Excel 自带筛选菜单。

没有匹配值时生成经当前列完整值集合确认不存在的唯一条件，并仍以 `xlFilterValues` 得到原生零行结果。所有值都匹配时不提交巨大数组，只清除目标字段自身条件。应用和清除均只调用目标 Field 的 `Range.AutoFilter`，不使用 `ShowAllData`、不关闭 `AutoFilterMode`，因此其他字段已有条件保持不变。

`BatchFilterService` 仅在内存中记录本工具最近应用的工作表、筛选标题行和目标绝对列，用于限制“清除本次过滤”只能清除本工具处理过的目标字段；不再保存隐藏行状态。批量操作临时调整 `ScreenUpdating` 和 `EnableEvents` 时，必须在 `finally` 中恢复原值。

### 最小职责划分

- `RibbonTest`：获取/校验 Excel 上下文并打开批量过滤窗体；不再包含早期测试事件。
- `BatchFilterForm`：处理输入、模式选择、提示和结果展示。
- `BatchFilterService`：识别/建立原生筛选范围、标准化条件、按 Equals/NotEquals/Contains/NotContains 计算匹配值集合、应用或清除目标字段 AutoFilter 及统计结果。
