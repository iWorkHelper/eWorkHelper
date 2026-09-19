# eWorkHelper 开发说明

## 开发流程

1. 阅读 `README.md` 和 `/docs` 中与任务相关的长期文档。
2. 先理解现有实现、VSTO 生命周期、Ribbon Designer 资源关系和 Excel Object Model 行为。
3. 采用最小修改，避免无关重构、公共接口变化或发布环境要求升级。
4. 修改后执行覆盖本次变更的构建、测试或人工验证。
5. 涉及功能、版本、发布或已知限制时同步更新文档。

## 文档导航

- [ARCHITECTURE.md](ARCHITECTURE.md)：技术基线、模块职责、Excel/VSTO 交互、签名与部署原则。
- [FEATURES.md](FEATURES.md)：批量过滤、取消合并并填充、关于入口和已知限制。
- [RELEASE.md](RELEASE.md)：版本规则、构建命令、发布前检查和当前变更摘要。
- [TESTING.md](TESTING.md)：发布前必须执行的人工回归清单。
- [CHANGELOG.md](CHANGELOG.md)：长期变更记录（唯一权威发布历史）。
- [REVIEW_TRACKING.md](REVIEW_TRACKING.md)：代码审查发现（E-01…E-18）的修复与验证记录。

## 当前项目基线

- 语言：C#。
- 目标框架：.NET Framework 4.8。
- 宿主：Microsoft Excel 桌面版。
- UI：VSTO Ribbon Designer + Windows Forms。
- 当前产品版本：`1.2.1`。

## 构建与调试

```powershell
msbuild .\eWorkhelper.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU" `
  /p:SignManifests=true /p:ManifestCertificateThumbprint=<thumbprint>
```

F5 调试需要 Visual Studio 2022、VSTO 工具、.NET Framework 4.8 Developer Pack 和本机 Excel。
VSTO manifest 构建需要签名证书（Office.targets 的 `ManageCertificateStore` 任务始终要求指纹），
可通过 `/p:ManifestCertificateThumbprint=<thumbprint>` 或 `IWORKHELPER_MANIFEST_CERT_THUMBPRINT`
环境变量提供；证书、私钥、指纹和用户级发布配置不得提交。

## 验证要求

- 发布构建目标为 0 Error / 0 Warning。
- 每次发布前按 [`TESTING.md`](TESTING.md) 完成人工回归，覆盖批量过滤（表格 / 已有 AutoFilter /
  普通标题行选择 / 四种匹配方式 / 显示文本语义 / 清除本次过滤 / 进度与取消）和取消合并并填充
  （单区域 / 多区域 / 公式 / 空值 / 错误值 / 数组公式 / 受保护工作表 / 状态恢复 / COM 释放）。
- 发布前必须扫描敏感信息、本机路径、构建产物、日志和临时文件。
