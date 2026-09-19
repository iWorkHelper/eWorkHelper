# eWorkHelper 发布说明

## 当前版本

当前产品版本：`1.3.0`。

需要四段数值版本的字段使用 `1.3.0.0`；历史 CLR/VSTO 兼容字段如无绑定需求可保持稳定值。

## 版本规则

- 对外产品版本使用 `Major.Minor.Patch`。
- `AssemblyInformationalVersion` 是 UI、Release Notes、问题反馈和发布说明使用的产品版本来源。
- `AssemblyVersion`、`AssemblyFileVersion`、`ApplicationVersion` 等必须满足宿主格式要求。
- UI 不硬编码版本号，必须从程序集元数据读取。

## 构建

VSTO 构建需要清单签名证书，必须通过 MSBuild 属性或 `IWORKHELPER_MANIFEST_CERT_THUMBPRINT`
环境变量提供证书指纹，否则会报 `MSB4044: 未给任务"ManageCertificateStore"的必需参数"CertificateThumbprint"赋值`。

```powershell
msbuild .\eWorkhelper.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU" `
  /p:SignManifests=true /p:ManifestCertificateThumbprint=<thumbprint>
```

或使用环境变量（`SignManifests` 会因指纹存在而自动为 `true`）：

```powershell
$env:IWORKHELPER_MANIFEST_CERT_THUMBPRINT = '<thumbprint>'
msbuild .\eWorkhelper.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU"
```

不要提交证书、私钥、指纹或用户级发布配置。

## 发布前检查

- 敏感信息：无 API Key、Token、密码、证书、私钥、本机绝对路径或真实用户数据。
- 产物：`bin/`、`obj/`、`publish/`、`.vsto`、`.manifest` 等不得进入 Git。
- 验证：Release 构建应达到 0 Error / 0 Warning，并按 [`TESTING.md`](TESTING.md) 完成 Excel 人工回归。
- GitHub 操作：提交、标签、推送和 Release 只在完成本地审查后执行。

## 发布历史

发布历史以 [`CHANGELOG.md`](CHANGELOG.md) 为唯一权威来源；根目录 `RELEASE_NOTES.md`
只保留当前版本的对外摘要并指向该文件。

## 当前版本摘要

### v1.3.0（2026-09-20）

- 集成 Excel 差异对比功能，支持多 Sheet 映射、差异高亮、基准值批注、行状态和可取消进度监控。
- 完善差异对比窗口布局、状态展示和安全取消/清理流程。
- 核心自动化测试扩展并完成现有功能回归验证。

### v1.2.1（2026-09-19）

- 更新 eWorkHelper 产品版本至 `1.2.1`。

### v1.2.0（2026-08-22）

- 统一当前发布版本为 `1.2.0`。
- 保留批量过滤、取消合并并填充、关于版本展示等当前功能。
- 精简 README 和 docs，保留长期维护信息。
- 修复代码审查发现（COM 释放、显示文本匹配、状态作用域、异常保护、进度与取消、文档漂移等），
  详见工作区根目录 `CODE_REVIEW_2026-09-13.md` 与 `docs/REVIEW_TRACKING.md`。
