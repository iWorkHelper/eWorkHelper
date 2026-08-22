# GitHub 发布检查清单

## 发布前安全门槛

任何 Git 初始化、Commit、Push、Merge、Tag 或 Release 前必须重新执行敏感信息检查。检查至少覆盖：

- Git tracked/untracked 文件、暂存区、差异和完整历史。
- 源码、注释、文档、配置、项目/解决方案、Designer、Resource、脚本和示例数据。
- 文件内容、文件名、目录名和路径。
- API Key、Token、密码、Secret、连接串、环境变量值、私钥、证书、签名凭据。
- 个人、公司、客户、部门、内部项目、邮箱、电话、用户名、机器名、内部域名/IP 和本地绝对路径。
- Visual Studio 缓存、日志、构建输出、Publish Profile、ClickOnce 输出和无关二进制文件。

任何无法确认安全的内容都必须先处理并重新扫描；未通过时禁止 Commit、Push 和 Release。`.gitignore` 不能替代对已跟踪文件和 Git 历史的检查。

## 构建与回归

- Debug Rebuild 通过。
- Release Rebuild 通过。
- Error 为 0，并审查全部 Warning。
- VSTO 引用、ThisAddIn、Ribbon Designer、资源与启动入口正常。
- 按 [`../Testing/RegressionChecklist.md`](../Testing/RegressionChecklist.md) 完成适用的自动和人工验证。
- 生成物只用于验证，不提交 `bin`、`obj`、DLL、PDB 或 Debug 输出。

## 版本与发布文档

- 按 [`../Versioning.md`](../Versioning.md) 确定完整四段产品版本。
- `AssemblyInformationalVersion`、Release Notes、Git Tag 和 GitHub Release 版本一致。
- Release Notes 只描述已实现且已验证的功能、要求、限制和安全说明。
- README 与当前功能、安装状态和文档入口一致。

## Git 与 GitHub

- 确认当前正式分支和待提交文件列表。
- 确认 remote 指向预期仓库，不猜测或改写 GitHub 用户名/地址。
- 不在命令、脚本或配置文件中保存 GitHub Token。
- Push 成功且认证有效后才创建 Tag 和 GitHub Release。
- 没有适合公开分发的安装包时，只发布源码和 Release Notes，不上传 Debug DLL、PDB、`bin` 或 `obj`。

## 发布记录

记录版本、日期、Commit ID、Tag、remote、Push 和 GitHub Release 状态。安全扫描报告只描述类型、位置和处理结果，不复述真实敏感值。
