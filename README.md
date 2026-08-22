# eWorkHelper

eWorkHelper 是一个面向 Microsoft Excel 桌面版的 VSTO 加载项。目前提供批量文本过滤工具：从当前单元格确定目标列，将多行条件转换为 Excel 原生 AutoFilter 条件，并尽量保留工作表中其他字段已有的筛选状态。

## 主要功能

- 在 Excel Ribbon 的 `eWorkHelper` 选项卡中提供“批量过滤”入口。
- 支持每行一个过滤条件，并自动去除空行、首尾空格和完全重复的条件。
- 支持“等于”“不等于”“包含”“不包含”四种不区分大小写的文本匹配方式。
- 支持 Excel 表格（ListObject）、已有 AutoFilter 区域，以及由用户选择标题行后建立的普通数据区域。
- 使用 Excel 原生 AutoFilter，不写改单元格、不创建辅助列或临时工作表。
- “清除本次过滤”只针对本工具最近处理的目标字段，不主动清除其他字段的筛选条件。

## 技术栈

- C#
- .NET Framework 4.8
- Visual Studio Tools for Office（VSTO）
- Excel Object Model / Office Interop
- Visual Studio 2022
- Windows Forms

## 运行环境要求

- Windows。
- Microsoft Excel 桌面版。
- Microsoft Visual Studio Tools for Office Runtime。
- 运行位数应与所使用的 Office/VSTO 环境兼容。

本仓库当前提供源代码，不包含预构建安装包或签名发布包。

## 开发环境要求

- Visual Studio 2022。
- 安装“.NET 桌面开发”和“Office/SharePoint 开发”相关工作负载或等效 VSTO 工具。
- .NET Framework 4.8 Developer Pack。
- 本机安装 Microsoft Excel 桌面版，用于 F5 调试和人工功能验证。

## 项目结构

```text
eWorkhelper.sln              Visual Studio 解决方案
eWorkhelper.csproj           Excel VSTO Add-in 项目
ThisAddIn.*                  VSTO Add-in 主机项和生成代码
RibbonTest.*                 Ribbon Designer、事件入口和资源
BatchFilterForm.cs           批量过滤窗口
BatchFilterService.cs        范围识别、条件匹配和 AutoFilter 逻辑
Properties/                  程序集、资源和设置
docs/                        开发约束、开发记录和安全审计
```

## 编译

1. 在 Visual Studio 2022 中打开 `eWorkhelper.sln`。
2. 确认 VSTO、.NET Framework 4.8 和 Excel 开发组件可用。
3. 选择 `Debug | Any CPU` 或 `Release | Any CPU`。
4. 选择“生成解决方案”或在 Developer PowerShell 中使用对应 Visual Studio MSBuild：

   ```powershell
   msbuild .\eWorkhelper.sln /t:Rebuild /p:Configuration=Debug /p:Platform="Any CPU"
   msbuild .\eWorkhelper.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU"
   ```

公开项目文件默认不绑定开发者证书。需要制作签名部署包时，请仅在受保护的本地或发布环境中配置证书，不要提交证书、私钥、指纹或用户级发布配置。

## 安装与使用

### 开发调试

1. 在 Visual Studio 中将 `eWorkhelper` 设为启动项目。
2. 选择 `Debug | Any CPU` 并按 F5；VSTO 项目将启动新的 Excel 实例。
3. 在 Excel 中打开工作簿，选择需要过滤列中的一个单元格。
4. 打开 `eWorkHelper` 选项卡，点击“批量过滤”。
5. 每行输入一个条件，选择匹配方式，然后点击“应用过滤”。

如果当前工作表没有 AutoFilter，工具会要求选择数据标题行。取消选择不会修改工作表。

### 分发安装

仓库尚未提供可直接安装的签名 ClickOnce/MSI 包。维护者需要在独立、受保护的发布环境中完成签名、信任和部署配置后再分发。

## 已知限制

- 仅支持 Windows 上的 Microsoft Excel 桌面版，不支持 Excel 网页版或 macOS 版。
- 当前界面和提示文本为中文。
- 匹配以单元格转换后的显示文本语义进行，不区分大小写；不提供正则表达式或公式条件。
- 工具依赖 Excel 原生 AutoFilter 的值列表能力；极大数据集或大量唯一值的性能取决于 Excel 和本机环境。
- “清除本次过滤”的状态保存在当前 Add-in 进程内存中，Add-in 卸载后不会保留该状态。
- 发布包制作、签名安装与不同 Office 版本/位数组合仍需维护者在目标环境中验证。

## 开发与贡献

提交修改前请先阅读 [`docs/DevelopmentConstraints.md`](docs/DevelopmentConstraints.md) 和 [`docs/Development.md`](docs/Development.md)。变更应保持 VSTO 生命周期、Ribbon Designer 资源关系和 Excel 原生筛选行为，并同时更新开发记录。建议至少完成 Debug、Release 构建及受影响 Excel 交互的人工验证。

## 安全与隐私

项目不应包含 API Key、Token、密码、证书、私钥、签名凭据、个人或组织身份信息、内部地址及本机绝对路径。公开发布前的审计摘要见 [`docs/SecurityReview.md`](docs/SecurityReview.md)。批量过滤逻辑在本地 Excel/VSTO 进程中运行；当前代码没有网络上传功能。

## License

本项目采用 [MIT License](LICENSE)。
