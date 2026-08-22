# eWorkHelper 开发日志

## 2026-08-22：GitHub 首次公开发布准备

### 任务目标

- 完成首次公开发布前的全仓库敏感信息审计、必要脱敏、公开文档与许可证整理。
- 在最终安全扫描通过后重新执行 Debug/Release 构建，并准备 Git 提交、推送及 GitHub Release。

### 安全要求

- API Key、Token、密码、私钥、证书、签名文件、证书指纹及其他认证材料不得进入公开仓库。
- 公司、部门、个人、邮箱、用户名、机器名、内部地址及含身份信息的本地路径必须移除或替换为中性内容。
- `.gitignore` 不能代替 Git 已跟踪内容和历史审计；任何安全检查未通过时禁止 Commit、Push 和创建 Release。
- 审计与记录只描述敏感信息类型、位置和处理结果，不复述真实敏感值。

### 检查范围

- `/docs` 开发文档、根目录文档、源码、注释、Designer/Resource、解决方案、项目文件、程序集元数据和配置。
- 文件与目录名称、构建/发布配置、ClickOnce/VSTO 签名设置、证书或密钥文件、测试数据、脚本及构建输出。
- Git tracked/untracked 内容、差异、远端配置和历史；完成脱敏及公开文档后进行二次与最终扫描。

### 当前状态

- 审计开始时目录尚未初始化为 Git 工作区，因此不存在既有分支、remote、tracked 内容或可供清理的 Git 历史；完成初次审计后已初始化新的本地仓库，仍无既有提交历史和 remote。
- 当前项目程序集版本为 `1.0.0.0`；后续 Release 版本以该既有版本为依据，不无故改动程序集版本。
- 初始状态不存在 README 和 App.config；项目原有本地开发签名绑定已在安全审计阶段移除。

### 公开发布文件

- 创建 `.gitignore`，排除 Visual Studio 缓存、Debug/Release 输出、`bin`、`obj`、测试/日志/临时文件、NuGet 本地产物、ClickOnce 发布输出、用户级 Publish Profile 以及证书和私钥文件。
- 创建 `README.md`，按实际代码说明批量过滤功能、技术栈、环境要求、项目结构、编译/使用方式、限制、贡献、安全隐私和许可证。
- 创建中性版权主体的 MIT `LICENSE`。
- 创建 `RELEASE_NOTES.md`，首次公开版本为 `v1.0.0`，与程序集版本 `1.0.0.0` 对齐。
- 创建 `docs/SecurityReview.md`，仅记录敏感信息类型、位置、脱敏措施、Git 历史状态和扫描结论。

### 敏感信息审计与脱敏

- 检查源码、文档、项目/解决方案、程序集元数据、Designer/Resource、配置、文件名、生成内容、VSTO/ClickOnce 设置和全部 Git 候选文件。
- 将程序集中的组织身份改为中性的 `eWorkHelper Contributors`。
- 从项目文件移除开发证书指纹，公开源码默认关闭清单签名；证书、私钥、指纹和用户级发布配置均不进入 Git。
- `.vs`、`bin`、`obj` 中的本机信息与构建产物已由 `.gitignore` 排除，且正常源码、项目、Designer 和 Resource 文件未被误忽略。
- 审计开始时没有 `.git`，因此不存在受污染的既有 Git 历史；新仓库在提交前仍需最终扫描。
- 文档创建后的二次扫描覆盖 23 个 Git 候选文件，未发现身份信息、邮箱、绝对路径、Secret、私钥、证书材料、证书指纹或发布端点。系统未安装 gitleaks，未为本任务引入额外依赖。

### Build 验证

- 公开项目默认不绑定证书；直接完整构建会在 VSTO 清单阶段要求本地签名配置，这是预期的发布环境约束。
- 使用当前用户证书存储中的本地开发证书，通过一次性 MSBuild 进程参数启用清单签名；证书标识未输出、未写入项目且不进入 Git。
- Visual Studio 2022 MSBuild 17.14：Debug / Any CPU Rebuild 通过，Error 0，Warning 0。
- Visual Studio 2022 MSBuild 17.14：Release / Any CPU Rebuild 通过，Error 0，Warning 0。
- VSTO 项目引用、Ribbon Designer、资源和生成代码均正常；未发现脱敏或重命名造成的缺失文件、路径或引用错误。
- 本次未执行 Excel GUI/F5 人工交互验证；既有人工验证基线保留，公开版本仍需在目标 Excel 环境验证安装、加载和筛选交互。

### Release 版本

- 最终准备版本：`v1.0.0`。
- Commit、Push、Tag 和 GitHub Release 只有在最终安全检查全部通过后才允许执行。

## 2026-08-22：批量过滤新增“不包含”匹配方式

### 本次需求与修改文件

- 在现有“等于”“不等于”“包含”基础上新增“不包含”，默认仍为“包含”，不改动其他功能。
- `BatchFilterService.cs`：在 `BatchFilterMatchMode` 增加 `NotContains` 并扩展内存匹配分支。
- `BatchFilterForm.cs`：新增互斥的“不包含”RadioButton，并接入 `NotContains` 枚举；原布局和其他控件保持不变。
- `docs/Development.md`：更新为四种匹配方式并明确“不包含”的多条件语义。
- `docs/DevelopmentLog.md`：记录本次修改和验证结果。

### 匹配逻辑与 AutoFilter

- `NotContains` 对每个输入条件使用不区分大小写的 `IndexOf` 检查；只要命中任一条件立即返回不匹配，遍历完所有条件均未命中才保留。因此逻辑为 `!Contains(c1) AND !Contains(c2) ...`，与现有“包含任一条件”完全互补。
- 条件 Trim、空行忽略、完全重复项去除、单次 `Value2` 读取、值转文本、完整匹配值去重均保持不变。
- 最终值集合仍通过目标字段 `Range.AutoFilter` 和 `xlFilterValues` 应用；ListObject、已有筛选范围、标题行选择、其他字段筛选保护和单字段清除未修改。
- 未引入 Hidden、辅助列、单元格写入或临时工作表。

### 编译与自动验证结果

- Visual Studio 2022 MSBuild 17.14 Debug / Any CPU Rebuild 通过，退出代码 0，生成 `bin/Debug/eWorkhelper.dll`。
- 指定样例通过：“包含”得到北京、北京市、北京科技有限公司、上海科技；“不包含”得到广州制造有限公司、深圳贸易有限公司。
- 逐项确认指定样例中 Contains 与 NotContains 完全互补。
- 单条件、多条件、重复条件、空行、空单元格、英文大小写和数字文本验证通过。
- 静态检查确认四个 RadioButton、默认 Contains、`NotContains` 分支和原生 AutoFilter 路径存在，且没有禁用的手动隐藏实现。

### Excel F5 实测结果与遗留问题

- 当前仍存在修改前启动且未结束的 Excel 调试进程，并打开用户现有工作簿；Visual Studio 显示正在运行。该进程加载的是旧程序集，无法用于验证本次新 DLL。
- 为避免强制关闭 Excel 导致未保存数据丢失，本次未终止该会话。因此新构建的四选项 UI 和实际 NotContains AutoFilter 尚未完成 F5 GUI 实测，不能声明 F5 验收通过。
- 待用户保存并关闭当前 Excel 后重新 F5 验证：四种模式互斥、默认包含、NotContains 实际筛选、其他字段条件保留和清除当前字段。

## 2026-08-22：调整批量过滤匹配方式并删除测试功能

### 本次需求

- 删除“包含全部条件”和“包含任意条件”模式，改为互斥的“等于”“不等于”“包含”，默认“包含”。
- 保留多行输入和“C# 内存匹配 + Excel 原生 AutoFilter”实现。
- 删除早期用于 VSTO 环境验证的 Ribbon 测试 Group、按钮、事件和提示框。

### 修改文件

- `BatchFilterService.cs`：将 `BatchFilterMode.Any/All` 改为 `BatchFilterMatchMode.Equals/NotEquals/Contains`，替换匹配判断。
- `BatchFilterForm.cs`：将两个旧 RadioButton 替换为“等于”“不等于”“包含”三个互斥选项，默认选中“包含”。
- `RibbonTest.cs`：删除 `btnTest_Click` 及测试 MessageBox，仅保留正式批量过滤入口。
- `RibbonTest.Designer.cs`：删除测试 Group、`btnTest`、事件绑定及对应 Designer 字段，只保留“数据工具 → 批量过滤”。
- `docs/Development.md`：清理测试按钮作为当前 UI 的描述，更新三种匹配方式设计。
- `docs/DevelopmentLog.md`：记录本次修改和验证结果。

### 三种匹配方式

- 等于：使用 `StringComparison.OrdinalIgnoreCase`，完整文本与任一条件相等即匹配，多个条件为 OR。
- 不等于：完整文本与所有条件都不相等才匹配，用于排除输入列表中的所有完整值。
- 包含：文本包含任一条件即匹配，多个条件为 OR；不再提供同时包含全部条件。
- 数据仍通过单次 `Value2` 读取并转换为文本；空值、数字及其他 Excel 值不会写回工作表。
- 匹配后的完整值集合继续交给目标字段的原生 `Range.AutoFilter(..., xlFilterValues, ...)`，范围识别、ListObject、标题行选择、其他字段筛选保护和单字段清除逻辑未改变。

### 自动验证与编译结果

- 指定样例验证通过：等于得到“北京、上海”；不等于得到“北京市、上海公司、深圳”；包含得到“北京、北京市、北京科技有限公司、上海科技”。
- 额外验证通过：单/多条件、重复项、空行、英文大小写、数字文本和空单元格判断。
- 源码清理检查通过：正式源码中不存在旧模式枚举、旧选项文本、`btnTest`、`grpTest`、`btnTest_Click` 或 `eWorkHelper 测试成功！`。
- 原生筛选路径静态检查通过：`InputBox(Type: 8)`、`Range.AutoFilter`、`xlFilterValues` 和单字段清除仍存在。
- Visual Studio 2022 MSBuild 17.14 Debug / Any CPU Rebuild 通过，退出代码 0，生成 `bin/Debug/eWorkhelper.dll`。

### Excel F5 实测结果与遗留问题

- 当前 Visual Studio 中存在修改前启动且未结束的 Excel 调试会话。Excel 窗口随后可见，但加载的是旧程序集，Ribbon 仍显示旧测试 Group，因此不能作为本次新构建的验收结果。
- 该 Excel 正打开用户现有工作簿；为避免强制终止进程造成未保存数据风险，本次未关闭该会话并重新启动。待用户保存并关闭当前 Excel 后，需要重新 F5 验证：测试 Group/按钮消失、三种 RadioButton、默认“包含”、实际 AutoFilter、其他字段保护和清除当前字段。
- 因此本次项目编译和自动逻辑验证通过，但“新构建 F5 Excel 实测通过”仍为待验证，不作成功声明。

## 2026-08-22：批量过滤改为 Excel 原生 AutoFilter

### 本次需求

- 保留现有批量过滤 Ribbon、窗体和多条件 AND/OR 交互，只替换底层行显示机制。
- 删除手动 Hidden 过滤及隐藏行恢复状态，最终必须由 Excel 原生 AutoFilter 控制筛选结果并保留筛选箭头。

### 修改文件

- `BatchFilterService.cs`：重写范围识别、标题行选择、匹配值集合、原生筛选和单字段清除实现。
- `BatchFilterForm.cs`：复用现有 UI，只将清除调用和无状态提示调整为目标 AutoFilter 字段语义。
- `RibbonTest.cs`：用户取消 Excel 标题行选择时静默结束，不弹出空错误提示；原按钮及窗体入口保持不变。
- `docs/Development.md`：改写批量过滤设计为“C# 内存条件计算 + Excel 原生 AutoFilter”。
- `docs/DevelopmentLog.md`：记录本次重构及验证状态。

### 移除的旧实现

- 删除 `FilterState` 隐藏行号缓存、可见行扫描、连续行区间隐藏和恢复逻辑。
- 批量过滤模块不再调用 `Hidden`、`EntireRow.Hidden` 或其他手动显示/隐藏行 API。
- 不使用 `ShowAllData`、`AutoFilterMode = false`、辅助列、临时工作表、单元格写值、公式或 VBA。

### AutoFilter 区域与标题行流程

- 启动时记录最初 `ActiveCell.Column`。
- 当前单元格位于 ListObject 时复用 `ListObject.Range` 和对应 `ListColumn.DataBodyRange`，并确保表格原生筛选箭头可见。
- 普通工作表已有 AutoFilter 时直接复用 `Worksheet.AutoFilter.Range`；目标列不属于该范围时按需求提示并停止，不重建范围。
- 没有 AutoFilter 时调用 `Application.InputBox(Type: 8)`，提示用户在 Excel 中选择标题行。取消时不修改工作表；跨多行时提示“请选择单一标题行”。
- 所选单元格的 `CurrentRegion` 只用于确定左右边界和底部；新 AutoFilter Range 的第一行严格等于用户所选标题行。最初目标列不在最终范围时停止。

### 原生筛选实现与已有筛选保护

- 目标数据列通过单次 `Value2` 读入内存，继续使用不区分大小写的 Contains AND/OR 匹配。
- 匹配后的完整值在内存中去重，以值数组作为 `Criteria1`、`xlFilterValues` 作为 Operator 调用目标 `Range.AutoFilter`。
- 无匹配时动态生成 GUID 形式、并经当前完整值集合确认不存在的条件，仍由 AutoFilter 产生零行结果。
- 全部匹配时省略 Criteria1，仅清除当前 Field 的条件；其他字段条件不变。
- 普通应用和“清除本次过滤”都只调用目标 Field。服务只记录本工具最近应用的工作表、标题行和目标绝对列，不保存行状态；清除不会关闭筛选箭头或影响其他 Field。
- 批量操作继续在 `finally` 中恢复原始 `ScreenUpdating` 和 `EnableEvents`。

### 编译与自动验证结果

- Visual Studio 2022 MSBuild 17.14：Debug / Any CPU `Rebuild` 通过，退出代码 0，生成 `bin/Debug/eWorkhelper.dll`。
- 反射逻辑验证通过：条件 Trim、空行/重复项处理，以及三个条件的 OR、AND 和 AND 反例。
- 禁用实现扫描通过：批量过滤相关源码中不存在 `Hidden`、`EntireRow`、`ShowAllData`、`AutoFilterMode =`、单元格 `Value2` 写入或新增 Sheet。
- 静态确认原生路径包含 `Application.InputBox(Type: 8)`、`Range.AutoFilter`、完整值数组和 `xlFilterValues`。

### Excel F5 实测结果与未验证项

- 已从当前打开的 Visual Studio 2022 实际按 F5。调试输出显示 `excel.exe` 启动并加载 VSTO/Interop 及 `eWorkhelper.dll`，随后 Excel 进程立即退出；当前未留下可操作的 Excel 窗口。因此本轮 F5 GUI 验收未通过，不能声明原生筛选运行场景已实测成功。
- 待人工验证：已有 AutoFilter 复用、无 AutoFilter 标题行选择和取消、标题上方有内容、筛选箭头、ListObject、三条件以上 AND/OR、无匹配、全部匹配、已有其他字段筛选及单字段清除。
- 已知边界：Excel `xlFilterValues` 自身对可选择唯一值数量存在宿主版本限制；插件未通过 Hidden 或辅助列绕过该原生限制。

## 2026-08-22：新增批量文本过滤功能

### 本次需求

- 在现有 `eWorkHelper` Ribbon 中新增“数据工具”Group 和“批量过滤”按钮。
- 支持按活动单元格所在列，对 ListObject 或 CurrentRegion 数据区执行任意数量文本条件的 AND/OR 包含过滤。
- 过滤不得修改数据，并可只恢复本功能造成的隐藏行。

### 新增文件

- `BatchFilterForm.cs`：批量过滤 WinForms 窗体。
- `BatchFilterService.cs`：范围识别、条件处理、匹配、隐藏与恢复服务。

### 修改文件

- `RibbonTest.cs`：保留测试按钮，新增批量过滤入口和服务实例。
- `RibbonTest.Designer.cs`：新增“数据工具”Group 与 `btnBatchFilter`。
- `eWorkhelper.csproj`：纳入新增 C# 文件。
- `docs/Development.md`：在代码修改前补充功能设计、数据范围、匹配和恢复方案。
- `docs/DevelopmentLog.md`：记录本次开发与验证结果。

### Ribbon 与 UI 实现

- 原有“测试”Group、`btnTest` 及提示行为保持不变。
- 新增“数据工具”Group 和“批量过滤”按钮。
- 窗体显示自动识别的列字母及可用表头，提供可滚动多行输入框、默认选中的“包含任意条件”、互斥的“包含全部条件”、应用、清除和关闭按钮。
- 空条件使用 MessageBox 提示；成功后只在窗体底部更新检查/匹配统计，不弹出成功消息框。

### 过滤实现方式

- 目标列来自 `ActiveCell.Column`。活动单元格位于 ListObject 时使用对应 `ListColumn.DataBodyRange`，否则使用 `CurrentRegion` 的首行以下数据。
- 输入按换行拆分，逐项 Trim，忽略空行并精确去除完全重复条件。
- 目标列 `Value2` 一次性读入内存；使用 `StringComparison.OrdinalIgnoreCase` 执行 Contains。任意模式逐条件 OR，全部模式逐条件 AND；支持空值、数字及任意合理数量条件。
- 只对应用前可见且不匹配的行进行隐藏；待处理行号合并为连续区间后设置 `EntireRow.Hidden`，避免逐单元格读取和逐行 COM 写入。
- 批量操作保存并在 `finally` 中恢复 `ScreenUpdating`、`EnableEvents`。

### 状态恢复方式

- `BatchFilterService` 在 Add-in 内存中保存本轮实际由工具隐藏的工作表引用和绝对行号，不向工作簿写入状态。
- 再次应用前先恢复上一轮工具隐藏行，再捕获当前可见行并重新过滤；“清除本次过滤”只恢复记录行。
- 应用前已经隐藏的行不会进入工具隐藏记录，因此不会被清除操作错误显示。

### 编译与验证结果

- Visual Studio 2022 MSBuild 17.14：`Rebuild`（Debug / Any CPU）通过，退出代码 0，生成 `bin/Debug/eWorkhelper.dll`。
- 反射逻辑验证通过：条件 Trim、空行忽略、重复项去除，以及需求示例中的 OR（北京科技、深圳科技）和 AND（仅北京科技）结果均符合预期。
- 静态确认新增文件已纳入项目，Ribbon 原测试控件未删除，新按钮事件已绑定。

### 实际 Excel 调试结果与未验证项目

- 本次当前自动化环境未执行 Excel F5/GUI 实测，不能声明 Ribbon 打开窗体、实际行隐藏和清除链路已经人工验证。
- 以下项目仍需在 Excel 中人工验证：单条件、3 个以上条件、中英文、数字/空单元格、无匹配/全部匹配、ListObject、普通 CurrentRegion、原有手工或 AutoFilter 隐藏行、连续多次应用、清除、保存/关闭工作簿边界，以及既有测试按钮基线。
- 当前没有已确认的运行时缺陷；内存状态随 Add-in 卸载而释放，因此若用户在过滤未清除时直接卸载 Add-in，本工具不会在卸载后自动恢复行。

## 2026-08-22：建立后续开发约束与确认基线

### 本次需求

- 将用户提供的 eWorkHelper 开发顺序、实现原则、编码约束、VSTO/Excel 约束、禁止事项、变更记录要求和完成标准固化为项目文档。
- 记录已经人工验证通过的 VSTO 基础链路，作为后续开发不得无故破坏的基线。

### 修改文件

- 新增 `docs/DevelopmentConstraints.md`。
- 更新 `docs/Development.md`，增加约束文档入口和已验证基线。
- 更新 `docs/DevelopmentLog.md`，记录本次文档变更。

### 关键实现方式

- 约束独立保存为长期维护文档，不修改 VSTO 项目、Ribbon 或业务代码。
- 后续任务开始前必须读取现有代码、项目结构以及 `/docs` 文档，并在完成后同步开发记录。

### 验证结果

- 用户人工验证：Visual Studio F5 启动 Excel、Add-in 加载、Ribbon 显示和测试按钮提示框均已通过。
- 本次仅修改 Markdown 文档，不改变编译产物或运行行为，因此未重复执行项目编译和 GUI 调试。

### 未验证项与后续事项

- 无新增运行时未验证项。
- 后续代码修改若影响既有 VSTO 基线，必须重新执行相应编译和人工调试验证。

## 2026-08-21：VSTO 基础链路验证

### 项目检查

- 确认为 Excel VSTO Add-in，而非其他 Office 插件架构。
- 目标框架为 .NET Framework 4.8，输出类型为 Library，根命名空间为 `eWorkhelper`。
- VSTO Host 为 Excel，LoadBehavior 为 3。
- Debug 使用 Office 16.0 Excel 注册表安装路径和 `/x` 参数，未发现自定义 Launcher。
- 已存在 VSTO、Office 和 Excel Interop 引用；`System.Windows.Forms` 已引用。
- 修改前不存在 Ribbon Designer 或 Ribbon XML。

### 实际修改

- 新增唯一一套 VSTO Ribbon Designer：`RibbonTest.cs`、`RibbonTest.Designer.cs`、`RibbonTest.resx`。
- 新增 `eWorkHelper` Tab、`测试` Group 和 `btnTest` 测试按钮。
- 将 `btnTest.Click` 绑定到 `btnTest_Click`。
- 使用 WinForms `MessageBox` 显示测试成功信息和“确定”按钮。
- 更新项目文件以包含 Ribbon Designer 的代码与资源。
- 修复原项目未启用 ClickOnce 清单签名而导致完整 VSTO 构建失败的问题：创建用户级开发代码签名证书，启用 `SignManifests` 并记录证书指纹。
- 新增开发说明和本日志。

### 测试结果

- 编译检查：通过。使用 Visual Studio 2022 MSBuild 17.14 执行 `Rebuild`（Debug / Any CPU），退出代码为 0，生成 `bin/Debug/eWorkhelper.dll` 及 VSTO/ClickOnce 清单。
- 静态检查：通过。确认 Ribbon Designer 的 `.cs`、`.Designer.cs`、`.resx` 均纳入项目；仅存在一套 Ribbon；命名空间均为 `eWorkhelper`；`btnTest.Click` 正确绑定到 `btnTest_Click`；界面标签和提示文字符合要求。
- Excel GUI/F5 验证：交付当时因当前 Visual Studio 管理员权限边界无法自动完成；用户随后于 2026-08-22 人工确认 F5 启动、Add-in 加载、Ribbon 显示和测试按钮提示框均通过。

### 发现的问题

- 当前目录不是 Git 工作区，因此无法使用 Git diff/status 跟踪本次文件变化。
- 原项目缺少 VSTO ClickOnce 清单签名设置；本次已为当前用户环境补齐。开发证书位于当前 Windows 用户证书存储，换机或换用户后需要重新选择开发证书。
- Visual Studio 已在本次外部修改前打开，人工验证前应在“检测到文件修改”对话框中选择“全部重新加载”。
