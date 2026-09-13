# eWorkHelper

eWorkHelper 是 [iWorkHelper](https://github.com/iWorkHelper) 旗下的 Microsoft Excel 桌面版 VSTO 加载项。它提供批量文本过滤与取消合并并填充工具，并尽量保持 Excel 原有数据、格式和其他字段筛选状态。

项目主页：<https://github.com/iWorkHelper/eWorkHelper>

## 当前版本

当前版本为 `1.2.0`。版本与发布规则见 [`docs/RELEASE.md`](docs/RELEASE.md)。

## 主要功能

- 在 Excel Ribbon 的 `eWorkHelper` 选项卡中提供“批量过滤”入口。
- 支持每行一个过滤条件，并自动去除空行、首尾空格和完全重复的条件。
- 支持“等于”“不等于”“包含”“不包含”四种不区分大小写的文本匹配方式。
- 支持 Excel 表格（ListObject）、已有 AutoFilter 区域，以及由用户选择标题行后建立的普通数据区域。
- 使用 Excel 原生 AutoFilter，不写改单元格、不创建辅助列或临时工作表。
- “清除本次过滤”只针对本工具最近处理的目标字段，不主动清除其他字段的筛选条件。
- 支持一次处理当前选区涉及的全部合并区域，取消合并后以原值或公式填充各单元格。
- Ribbon 末端提供“关于”按钮，显示插件简介和从程序集动态读取的完整版本号。

## 技术栈

- C#
- .NET Framework 4.8
- Visual Studio Tools for Office（VSTO）
- Excel Object Model / Office Interop
- Visual Studio 2022
- Windows Forms

## 运行环境

- Windows。
- Microsoft Excel 桌面版。
- Microsoft Visual Studio Tools for Office Runtime。
- 运行位数应与所使用的 Office/VSTO 环境兼容。

## 安装 / 使用

最新 Release：<https://github.com/iWorkHelper/eWorkHelper/releases/latest>

Release 中的编译组件包用于版本核验和受控部署，不是可直接双击安装的 ClickOnce/MSI 安装包。正式部署仍需维护者在受保护环境中使用受信任证书完成签名、信任和安装配置。

开发调试时：

1. 在 Visual Studio 中将 `eWorkhelper` 设为启动项目。
2. 选择 `Debug | Any CPU` 并按 F5；VSTO 项目将启动新的 Excel 实例。
3. 在 Excel 中打开工作簿，选择需要过滤列中的一个单元格。
4. 打开 `eWorkHelper` 选项卡，点击“批量过滤”。
5. 每行输入一个条件，选择匹配方式，然后点击“应用过滤”。

如需取消合并并填充，请选择一个或多个包含合并单元格的区域，然后点击 Ribbon 中的“取消合并并填充”。公式会按 Excel 的 R1C1 复制语义填充。如果当前工作表没有 AutoFilter，工具会要求选择数据标题行；取消选择不会修改工作表。

> `eWorkhelper` 是现有 Visual Studio 工程内部标识。为避免影响 VSTO、ClickOnce、调试和升级兼容，本轮不调整其大小写；面向用户的品牌名称统一为 `eWorkHelper`。

## 开发环境

- Visual Studio 2022。
- 安装“.NET 桌面开发”和“Office/SharePoint 开发”相关工作负载或等效 VSTO 工具。
- .NET Framework 4.8 Developer Pack。
- 本机安装 Microsoft Excel 桌面版，用于 F5 调试和人工功能验证。

### 项目结构

```text
eWorkhelper.sln              Visual Studio 解决方案
eWorkhelper.csproj           Excel VSTO Add-in 项目
ThisAddIn.*                  VSTO Add-in 主机项和生成代码
MainRibbon.*                 Ribbon Designer、事件入口和资源
BatchFilterForm.cs           批量过滤窗口
BatchFilterService.cs        范围识别、条件匹配和 AutoFilter 逻辑
UnmergeAndFillService.cs     合并区域发现、去重、取消合并与内容填充
Properties/                  程序集、资源和设置
docs/                        架构、功能、开发、发布和变更文档
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

## 已知限制

- 仅支持 Windows 上的 Microsoft Excel 桌面版，不支持 Excel 网页版或 macOS 版。
- 当前界面和提示文本为中文。
- 匹配以单元格转换后的显示文本语义进行，不区分大小写；不提供正则表达式或公式条件。
- 工具依赖 Excel 原生 AutoFilter 的值列表能力；极大数据集或大量唯一值的性能取决于 Excel 和本机环境。
- “清除本次过滤”的状态保存在当前 Add-in 进程内存中，Add-in 卸载后不会保留该状态。
- 发布包制作、签名安装与不同 Office 版本/位数组合仍需维护者在目标环境中验证。

## 开发与贡献

开发文档总入口为 [`docs/Development.md`](docs/Development.md)。涉及架构、功能或发布时同步阅读 [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md)、[`docs/FEATURES.md`](docs/FEATURES.md) 和 [`docs/RELEASE.md`](docs/RELEASE.md)。变更应保持 VSTO 生命周期、Ribbon Designer 资源关系和 Excel 原生筛选行为，并完成适用的构建及 Excel 人工回归。

## 安全与隐私

项目不应包含 API Key、Token、密码、证书、私钥、签名凭据、个人或组织身份信息、内部地址及本机绝对路径。公开发布前按 [`docs/RELEASE.md`](docs/RELEASE.md) 完成安全检查。批量过滤逻辑在本地 Excel/VSTO 进程中运行；当前代码没有网络上传功能。

## 文档入口

- [架构说明](docs/ARCHITECTURE.md)
- [功能说明](docs/FEATURES.md)
- [开发说明](docs/Development.md)
- [发布说明](docs/RELEASE.md)
- [变更记录](docs/CHANGELOG.md)
- [Release Notes](RELEASE_NOTES.md)

## Release

查看[最新 Release](https://github.com/iWorkHelper/eWorkHelper/releases/latest)或[全部 Releases](https://github.com/iWorkHelper/eWorkHelper/releases)。发布资产的用途和限制见“安装 / 使用”章节。

## License

本项目采用 [MIT License](LICENSE)。
