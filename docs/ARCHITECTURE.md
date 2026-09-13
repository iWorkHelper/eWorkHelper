# eWorkHelper 架构

## 项目定位

eWorkHelper 是 Excel 桌面版 VSTO 加载项，当前提供批量文本过滤、取消合并并填充、关于信息等本地数据工具。插件运行在 Excel/VSTO 进程内，不包含网络上传逻辑。

## 技术基线

- C# / .NET Framework 4.8
- Visual Studio Tools for Office (VSTO)
- Excel Object Model / Office Interop
- Windows Forms
- Visual Studio 2022

## 主要组件

- `ThisAddIn.*`：VSTO 加载项入口和宿主生命周期。
- `MainRibbon.*`：Excel Ribbon UI、按钮事件和关于窗口。
- `BatchFilterForm.cs`：批量过滤窗口。
- `BatchFilterService.cs`：筛选区域识别、条件匹配、AutoFilter 应用和清除。
- `UnmergeAndFillService.cs`：合并区域发现、取消合并、值或公式填充和状态恢复。
- `Properties/` 与 `Resources/`：程序集元数据、设置、资源和图标。

## Excel / VSTO 交互原则

- 优先使用 Excel 原生 AutoFilter，不改写数据、不创建辅助列或临时工作表。
- 普通区域需要用户确认标题行；用户取消时不得修改工作表。
- 取消合并并填充只处理当前选择触达的合并区域，并在异常路径恢复 Excel 交互状态。
- UI 版本展示从当前程序集 `AssemblyInformationalVersion` 动态读取，不维护硬编码显示版本。

## 签名与部署

VSTO manifest 构建需要签名证书。源码中不提交 PFX、私钥或证书指纹；本地/发布构建应通过受保护证书存储、环境变量或 MSBuild 参数提供证书。生产分发需要正式代码签名与目标 Office 环境验证。
