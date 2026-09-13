# eWorkHelper 发布说明

## 当前版本

当前产品版本：`1.2.0`。

需要四段数值版本的字段使用 `1.2.0.0`；历史 CLR/VSTO 兼容字段如无绑定需求可保持稳定值。

## 版本规则

- 对外产品版本使用 `Major.Minor.Patch`。
- `AssemblyInformationalVersion` 是 UI、Release Notes、问题反馈和发布说明使用的产品版本来源。
- `AssemblyVersion`、`AssemblyFileVersion`、`ApplicationVersion` 等必须满足宿主格式要求。
- UI 不硬编码版本号，必须从程序集元数据读取。

## 构建

```powershell
msbuild .\eWorkhelper.sln /t:Restore,Rebuild /p:Configuration=Release /p:Platform="Any CPU" /p:ManifestCertificateThumbprint=<thumbprint>
```

VSTO 构建需要清单签名证书。不要提交证书、私钥、指纹或用户级发布配置。

## 发布前检查

- 敏感信息：无 API Key、Token、密码、证书、私钥、本机绝对路径或真实用户数据。
- 产物：`bin/`、`obj/`、`publish/`、`.vsto`、`.manifest` 等不得进入 Git。
- 验证：Release 构建应达到 0 Error / 0 Warning，并完成 Excel 人工回归。
- GitHub 操作：提交、标签、推送和 Release 只在完成本地审查后执行。

## 变更摘要

### v1.2.0

- 统一当前发布版本为 `1.2.0`。
- 保留批量过滤、取消合并并填充、关于版本展示等当前功能。
- 精简 README 和 docs，保留长期维护信息。
