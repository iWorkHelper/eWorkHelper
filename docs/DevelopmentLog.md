# eWorkHelper 开发日志

本日志只保留重要功能、架构、兼容性、版本和发布准备变更。临时调试过程、已失效方案及逐次尝试不作为长期记录保存。

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
