# VSTO 技术架构

## 技术基线

- C#、.NET Framework 4.8。
- Visual Studio 2022 VSTO 项目系统。
- 宿主为 Windows 桌面版 Microsoft Excel。
- Office/Excel Interop 15.0 引用使用嵌入互操作类型。
- 项目输出类型为类库，VSTO LoadBehavior 为 `3`。

## 主要组件

- `ThisAddIn.*`：VSTO Add-in 主机项与生命周期入口。Startup 和 Shutdown 保持轻量，不自行创建 Excel 进程。
- `MainRibbon.*`：唯一的 Ribbon Designer 实现，提供 `eWorkHelper` Tab 和正式功能入口。
- `BatchFilterForm.cs`：批量过滤交互窗口，负责条件输入、模式选择、提示和结果展示。
- `BatchFilterService.cs`：负责 Excel 上下文识别、筛选区域选择、条件标准化、匹配计算与原生 AutoFilter 操作。
- `Properties/`：程序集元数据、Visual Studio 资源与设置基础设施。

## Excel / VSTO 交互

Ribbon 事件从 `Globals.ThisAddIn.Application` 获取当前 Excel Application，不创建额外 COM Application。服务基于当前 Workbook、Worksheet 和 ActiveCell 建立操作上下文，并明确限定目标列与筛选区域。

Excel 表格使用 `ListObject.Range`；已有普通 AutoFilter 时复用 `Worksheet.AutoFilter.Range`；未启用 AutoFilter 时由用户选择标题行，再依据 CurrentRegion 确定左右边界和数据底部。

批量过滤只通过目标字段的 `Range.AutoFilter` 应用或清除条件，不写改单元格，不创建辅助列或临时工作表，也不直接设置行 Hidden 状态。其他字段已有筛选条件应保持不变。

## 状态与性能原则

- 目标列 `Value2` 一次性读入内存，减少逐单元格 COM 调用。
- 匹配后的完整值集合去重后交给 Excel 原生 `xlFilterValues`。
- 工具仅在 Add-in 内存中记录最近处理的工作表、标题行和目标列，用于限制“清除本次过滤”的作用范围。
- 临时调整 `ScreenUpdating` 和 `EnableEvents` 时必须在 `finally` 中恢复原值。
- 避免长期持有无必要的 Excel COM 对象，不在 UI 线程执行无边界的耗时工作。

## 构建与签名

公开项目不绑定开发者证书，不保存证书指纹、私钥或签名凭据。完整 VSTO 清单构建需要维护者在受保护的本地或发布环境中通过临时配置提供签名证书；认证材料不得进入 Git。

Debug 使用 VSTO 原生 Excel Host/Start Action，依据 Office 注册表安装位置启动 Excel，并使用 `/x` 启动参数。
