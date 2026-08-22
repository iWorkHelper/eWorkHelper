# eWorkHelper 开发说明

本文档是 `/docs` 的开发入口。开始任何修改前，先阅读 [`DevelopmentConstraints.md`](DevelopmentConstraints.md)，再根据任务类型阅读对应设计、测试、版本或发布文档。

## 开发流程

1. 阅读开发约束、相关设计和现有代码。
2. 明确需求、影响范围、方案和风险；不确定用途的文件不得直接删除。
3. 只修改任务所需内容，保持 VSTO 生命周期、Excel 对象边界和现有业务行为。
4. 执行与风险相称的静态检查、自动测试、Debug/Release 构建和 Excel 人工验证。
5. 更新当前设计文档和精简的开发日志，明确未验证项与已知限制。
6. 进行 Git 操作或发布前，必须先完成敏感信息检查和发布清单。

## 文档导航

### 核心规范

- [`DevelopmentConstraints.md`](DevelopmentConstraints.md)：所有开发长期遵守的约束。
- [`Versioning.md`](Versioning.md)：唯一版本号规则及 .NET Framework 元数据映射。
- [`DevelopmentLog.md`](DevelopmentLog.md)：重要功能、架构、兼容性与发布准备的精简变更记录。

### 架构

- [`Architecture/VstoArchitecture.md`](Architecture/VstoArchitecture.md)：技术基线、模块职责和 Excel/VSTO 交互原则。

### 正式功能

- [`Features/BatchFiltering.md`](Features/BatchFiltering.md)：批量文本过滤的当前行为、范围规则和限制。

### 测试

- [`Testing/RegressionChecklist.md`](Testing/RegressionChecklist.md)：构建、VSTO 基线和批量过滤长期回归清单。

### 发布

- [`Release/ReleaseChecklist.md`](Release/ReleaseChecklist.md)：GitHub 发布、安全审计、构建和版本检查清单。

## 当前项目基线

- Visual Studio 2022、C#、.NET Framework 4.8、VSTO 4.0。
- Windows 桌面版 Microsoft Excel 宿主。
- 解决方案：`eWorkhelper.sln`；项目：`eWorkhelper.csproj`。
- 唯一 Ribbon 实现：`MainRibbon.cs`、`MainRibbon.Designer.cs`、`MainRibbon.resx`。
- 当前正式功能：批量文本过滤，使用 C# 内存匹配和 Excel 原生 AutoFilter。
- 当前完整产品版本：`1.0.260822.1`，由 `AssemblyInformationalVersion` 表示。
- 公开源码不绑定开发者证书，不包含签名凭据或预构建安装包。

## 已验证基线

历史人工验证已确认 Visual Studio F5 启动 Excel、Add-in 加载、Ribbon 显示和按钮事件链路可用。当前正式功能后续变更必须按回归清单重新验证受影响部分，不得把历史测试按钮或临时调试入口重新引入项目。

## 构建与调试

1. 使用 Visual Studio 2022 打开 `eWorkhelper.sln`。
2. 确认安装 .NET Framework 4.8 Developer Pack、VSTO/Office 开发工具和桌面版 Excel。
3. 分别选择 `Debug | Any CPU`、`Release | Any CPU` 构建。
4. F5 调试使用项目内 Excel Host/Start Action 启动新的 Excel 实例。

完整 VSTO 清单构建需要本地签名配置。证书、私钥、指纹和用户级发布配置只允许存在于受保护的本地/发布环境，不得提交到 Git。

## 当前项目结构

```text
eWorkhelper.sln
eWorkhelper.csproj
ThisAddIn.*
MainRibbon.*
BatchFilterForm.cs
BatchFilterService.cs
Properties/
docs/
```

项目没有独立自动化测试项目。任何新增长期有效测试应纳入解决方案并在回归清单中记录，不得使用临时测试入口替代可维护测试。
