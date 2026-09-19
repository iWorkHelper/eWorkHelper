# eWorkHelper 审查发现修复跟踪

来源：`CODE_REVIEW_2026-09-13.md`（工作区根目录审查报告）
建立日期：2026-09-13
提交策略：**不提交 git**，改动保留在工作树。

## 状态说明

| 状态 | 含义 |
|---|---|
| 待修复 | 尚未开始 |
| 修复中 | 正在处理 |
| 已修复 | 代码已改，尚未验证 |
| 已验证 | 已通过编译/运行验证并记录结论 |
| 不适用 | 经复核确认为非缺陷或需产品决策，已说明原因 |

## 验证方式

- 编译：`msbuild eWorkhelper.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU" /p:SignManifests=true /p:ManifestCertificateThumbprint=<thumbprint>`
- 需 0 Error。注意：`README.md` 的构建命令缺证书指纹会失败（见 E-10）。
- 运行期（Excel 人工回归）清单见 [`TESTING.md`](TESTING.md)。

**本次会话最终验证命令与结果（2026-09-13）**

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  eWorkhelper.sln /t:Rebuild /p:Configuration=Release /p:Platform="Any CPU" `
  /p:SignManifests=true /p:ManifestCertificateThumbprint=EC8D8202A516023590853D2852D978342575D594 /v:minimal /nologo
# EXITCODE=0
# eWorkhelper -> ...\bin\Release\eWorkhelper.dll
```

另外以 `/v:detailed` 抓取完整日志统计 `: warning ` 与 `: error `：Release 与 Debug 两个配置均为 **0 warning / 0 error / 退出码 0**。
环境变量路径也已实测：设置 `IWORKHELPER_MANIFEST_CERT_THUMBPRINT` 后不带 `/p:SignManifests`、不带
`/p:ManifestCertificateThumbprint` 构建同样退出码 0。

---

## 严重

### E-01 全项目零 COM 释放，存在不可回收的中间 RCW
- **位置**：`BatchFilterService.cs`、`UnmergeAndFillService.cs` 全文件；`MainRibbon.cs:17,37`
- **问题**：无任何 `Marshal.ReleaseComObject`/`FinalReleaseComObject`。大量双点写法产生无法释放的中间 RCW（`:257 ListColumns[f].DataBodyRange`、`:293 Cells[1,1].CurrentRegion`、`:330 Columns[f].Offset[1,0].Resize[...]`、`:309` 两处 `Cells`）。`UnmergeAndFillService` 递归每节点新建 5 个中间 RCW（`:111-112,119-120,145-148`）。
- **影响**：工作簿关闭后仍被引用、Excel 进程无法干净退出；批量操作 RCW 数暴涨。
- **方案**：引入 `ComHelper.Release(object)` 助手；每个双点链改局部变量并在 `finally` 中 `FinalReleaseComObject`；不得释放 `Globals.ThisAddIn.Application`。
- **实际处置**：
  1. 新增 `ComHelper.cs`（唯一公开成员 `Release(object)`）。`Release` 对 null / 非 COM 对象静默返回，并吞掉释放路径异常。
  2. `BatchFilterService.cs` 全量改造：`application.Workbooks`、`ActiveSheet`、`ActiveCell`、`ListObject`、`worksheet.AutoFilter`、`autoFilter.Range`、`listObject.Range`/`DataBodyRange`/`AutoFilter`/`ListColumns[f]`/`DataBodyRange`、`selectedRange.Rows`、`selectedRange.Cells[1,1]`、`CurrentRegion`、`region.Columns`/`Rows`、`worksheet.Cells[...]`、`worksheet.Range[...]`、`filterRange.Rows`/`Columns`、`Columns[f]`、`Offset[...]`、`Resize[...]`、`filterRange.Cells[1,f]`、`dataRange.Rows`、`dataRange.Text`/`Value2`、`SpecialCells(...)`、`Areas` 与每个 `area`、`AutoFilter.Filters`、`Range.Text`/`Value2` 返回的数组全部改为局部变量并在 `finally` 或紧邻作用域内逆序释放。
  3. `UnmergeAndFillService.cs` 全量改造：`application.Selection`、`selection.Areas` 与每个 `area`、`range.Rows`/`Columns`/`Cells[1,1]`、`MergeArea`、`Resize`、`Offset`、`mergeArea.Cells[1,1]`、`container.Rows`/`Columns` 全部释放；递归中发现并采纳一个**所有权转移**机制（`AddMergeArea` 返回 bool），只有未被 `mergeAreas` 接管的合并区域才在本地释放，避免重复释放与提前释放。
  4. `BatchFilterContext` 现在登记（`Track`）并在 `Dispose` 中逆序释放其持有的全部 RCW；`Globals.ThisAddIn.Application` **从不释放**（`ComHelper` 的 XML 注释与 `BatchFilterContext.Dispose` 注释均明确写出该约束）。
  5. 保留 `foreach (object value in array)` 这类托管数组遍历（非 COM 集合），未引入无意义的释放。
- **状态**：已验证
- **验证**：
  - 命令：上述最终验证命令 → **退出码 0**；`/v:detailed` 全日志 0 warning / 0 error（Release + Debug）。
  - 代码级证据：`Select-String -SimpleMatch 'ComHelper.Release'` 计数 —— `BatchFilterService.cs` = **52 处**，`UnmergeAndFillService.cs` = **19 处**；全仓 `Marshal.FinalReleaseComObject` 调用点唯一（`ComHelper.cs:33`，即助手内部一处）。
  - 运行期（Excel 进程能否干净退出、句柄是否持续增长）属真机验证范围，已列入 `docs/TESTING.md` 第 7 节。

## 高

### E-02 过滤匹配读原始 `Value2`，日期/百分比列隐藏整张表
- **位置**：`BatchFilterService.cs:178,185`；文档 `README.md:95`
- **方案**：改用显示文本语义（`Range.Text`/`DisplayFormat`）或传强类型值；保持与文档承诺一致。
- **实际处置**：
  1. 新增 `ReadCellDisplays(...)`：一次性读入目标列的 `Range.Text` 与 `Range.Value2`，逐行用 `ResolveDisplayText(text, raw)` 解析显示文本。
  2. `ResolveDisplayText` 以 `Range.Text` 为主；当列宽不足导致 `Range.Text` 返回纯 `#` 串（正则 `^#+$`）时回退到 `FallbackText(raw)`（字符串原样、`DateTime` 用 invariant `yyyy-MM-dd HH:mm:ss`、数值用 invariant `General` 文本、布尔用 `TRUE`/`FALSE`），避免把 `#####` 当成真实内容参与匹配与值列表。
  3. 匹配改在显示文本上进行（日期列按界面显示的日期文本匹配，不再是 OADate 序列号）。
  4. 百分比兼容：`BuildAlternateText` 为显示文本 `15%` 额外提供 `0.15`，为千分位等区域格式额外提供 invariant 原始数值文本（仅当与显示文本不同），`IsMatch(CellDisplay, ...)` 主匹配显示文本、失败后再匹配该辅助候选，因此用户输入 `15%` 或 `0.15` 都能命中。
  5. 传给 Excel 的 `xlFilterValues` 条件仍取自同一次 `Range.Text` 读取的字符串（小写/空值语义与原实现一致，空字符串映射为 `"="`），因此“内存判定”和“Excel 值列表”使用同一文本来源，不会出现判定与显示不一致。
- **状态**：已验证（编译与 0 warning 已验证；语义正确性需按 TESTING 第 3 节在真机 Excel 上确认）
- **验证**：命令同 E-01 → 退出码 0。`grep` 确认 `ReadCellDisplays` 内不再出现 `DataRange.Value2` 作为匹配来源，匹配文本只来自 `Range.Text`（`Value2` 仅用于 `#####` 溢出回退与百分比/区域格式辅助候选）；`Range.Text` 的 255 字符截断由“`#` 串回退 + 保留原 255+ 字符原文进行匹配”处理。
  - 诚实说明：`Range.Text` 在列宽不足时会返回 `#####`，本实现按 `^#+$` 识别并回退；但“AutoFilter 自身对超长/超宽文本的显示是否与 `Range.Text` 完全一致”无法在本无 Excel 会话的环境中实测，已列入 `docs/TESTING.md` 第 3 节请维护者确认。

### E-03 Workbook/Worksheet RCW 被加载项生命周期长期缓存
- **位置**：`BatchFilterService.cs:570-571,584-585,18-31`
- **方案**：`AppliedFilterState` 只存身份信息（`FullName`/`Name`/行列号）；`BatchFilterContext` 实现 `IDisposable`；`BatchFilterForm.Dispose` 释放。
- **实际处置**：`AppliedFilterState` 改为只保存 `workbookFullName` / `worksheetName` 字符串与 `filterRow`、`filterColumn`、`filterRowCount`、`filterColumnCount`、`headerRow`、`targetColumn`、`fieldIndex` 整数（`filterRow/Column` 在构造时一次性读出，行/列计数由 `CountOf` 即时读取并释放 `Rows`/`Columns`）。`Matches(...)` 全部改用字符串 + 整数比较，并对已关闭工作簿/工作表的 `COMException`（`RPC_E_DISCONNECTED`）与 `InvalidCastException` 做了捕获，按“不匹配”返回。`BatchFilterContext` 实现 `IDisposable`（登记-逆序释放 + `disposed` 幂等标志），`BatchFilterForm.Dispose(bool)` 调用 `context.Dispose()`，`MainRibbon` 在 `finally` 中 `form.Dispose()`，二者保证上下文必然释放且只释放一次。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0。`grep` 确认 `AppliedFilterState` 字段中不再有 `Excel.Workbook` / `Excel.Worksheet` 类型字段；`BatchFilterContext : IDisposable` 与 `protected override void Dispose(bool disposing)` 均已存在。

### E-04 `LoadInitialState` 与 Ribbon 回调缺异常保护
- **位置**：`BatchFilterForm.cs:137`、`BatchFilterService.cs:350-358`、`MainRibbon.cs:27-30`
- **方案**：初始化移出构造函数或加 try/catch；`btnBatchFilter_Click` 补 try/catch（对齐 `btnUnmergeAndFill_Click`）；`GetTargetFilter` 守护 `Filters` 访问。
- **实际处置**：
  1. `BatchFilterForm` 构造函数中读取初始状态的整段包入 `try/catch(Exception)`：失败时清空条件框、匹配模式回落 `Contains`、结果标签显示“无法读取当前筛选状态，请手工输入过滤条件。”，并弹出带原因的警告；窗体仍可用，不会出现 CLR 未处理异常对话框。
  2. `MainRibbon.btnBatchFilter_Click` 整体包入 `try/catch(Exception)`，失败弹出“批量过滤失败：…”（对齐 `btnUnmergeAndFill_Click` 的“操作失败：…”模式）；`finally` 中释放未被窗体接管的上下文。
  3. `GetTargetFilter` 把 `context.AutoFilter.Filters` 的访问放入 `try`，捕获 `COMException` / `InvalidCastException` 并返回 null，同时释放 `Filters` RCW；`FieldIndex < 1` 与 `null` 场景直接返回 null。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0；`/v:detailed` 0 warning / 0 error。代码级：`BatchFilterForm` 构造函数内 `LoadInitialState` 位于 `try` 块中；`btnBatchFilter_Click` 有顶层 `catch (Exception)`；`GetTargetFilter` 有 `catch (COMException)`。

### E-05 长任务阻塞 UI 线程，无进度、无取消
- **位置**：`BatchFilterForm.cs:152-188`、`BatchFilterService.cs:169-219`
- **方案**：`WaitCursor` + `StatusBar` + 分块进度 + 可取消；`FormClosing` 守护；恢复状态在 `finally`。
- **实际处置**：
  1. Excel COM 调用**全部留在 UI 线程**（未引入任何后台线程 / `Task` / `BackgroundWorker`）。
  2. `BatchFilterService.Apply` 在 `finally` 中逐项恢复：`Application.Cursor`（`xlWait` → 原值，读取失败则不恢复）、`Application.StatusBar = false`、`Application.Calculation`（操作期间 `xlCalculationManual` → 原值）、`EnableEvents`、`ScreenUpdating`。每个恢复动作**各自独立** `try/catch`（同时满足 E-08）。
  3. 进度：`ReadCellDisplays` 每 `rowCount/100` 行（最少 1 行）回调一次 `BatchFilterProgressReport`，文本形如“正在读取数据 1200/50000 … 按 Esc 取消”；`BatchFilterForm` 把它写入结果标签并调用 `Application.DoEvents()` 泵消息（每次回调前检查 `IsDisposed`）。
  4. 取消：`BatchFilterCancellationCheck` 委托在每个进度点被调用，为 true 时抛 `OperationCanceledException`；窗体持有 `CancellationTokenSource`，`KeyPreview = true` + `KeyDown` 在操作进行中把 Esc 映射为取消（操作结束后 Esc 仍交由 `CancelButton` 正常关闭）。取消后 `finally` 恢复全部 Excel 状态，提示“操作已取消，工作表筛选状态已恢复。”
  5. `FormClosing` 守护：`OnFormClosing` 在 `busy` 时 `e.Cancel = true` 并请求取消，避免 `DoEvents` 泵出的关闭消息把窗体释放掉后继续访问已释放对象；`busy` 期间禁用“应用过滤”“清除本次过滤”并直接忽略重复点击，杜绝重入。
  6. `BatchFilterForm.Dispose(bool)` 在 `busy` 时先 `Cancel()` 再释放上下文。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0；0 warning / 0 error。代码级：`Application.DoEvents` 仅出现在窗体进度回调中；`BatchFilterService.cs` 中无 `Thread`/`Task`/`BackgroundWorker`；`Application.Cursor`、`Application.StatusBar`、`Application.Calculation` 的恢复点均位于 `finally`。真机可见性验证列入 `docs/TESTING.md` 第 5 节。

## 中

### E-06 报告匹配数与实际显示可能不符
- **位置**：`BatchFilterService.cs:181-192` vs `:158`
- **方案**：统一可见性语义；或明示为“条件匹配数”而非“行数”。
- **实际处置**：`BatchFilterResult` 同时给出 `CheckedCount`、`MatchedRowCount`（仅目标列的条件匹配数）、`VisibleCount`（应用筛选后目标列实际可见行数，`-1` 表示未统计）。`Apply` 成功后调用 `CountVisibleDataRows` 用 `SpecialCells(xlCellTypeVisible)` + 逐 `area` 的 `Rows.Count` 统计真实可见行数（无可见单元格时返回 0，不抛原始 COM 错误）。窗体展示规则：两者相等时输出“共检查 N 行，匹配 M 行。”；不等时输出“共检查 N 行；本次条件匹配 M 行；当前实际显示 V 行（其它字段筛选会取交集）。”，把两种语义显式区分开。`LoadInitialState` 与 `Apply` 都改为读取同一套 `Range.Text` 显示文本，读写两条路径的可见性/取值语义现已一致。另修正原实现“0 匹配 → 传占位 GUID → 隐藏整张表”的行为：0 匹配时**不改动工作表**并明确提示（T-3 已列入清单）。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0。代码级：`BatchFilterResult.MatchedRowCount` 与 `VisibleCount` 分列；窗体两种文案分支存在；`ReadVisibleUniqueValues` 与 `ReadCellDisplays` 均走 `Range.Text`。

### E-07 失败后未回滚工作表变更
- **位置**：`BatchFilterService.cs:102-103,255`
- **方案**：先校验再建 AutoFilter，或失败时回滚 `AutoFilterMode`/`ShowAutoFilter`。
- **实际处置**：
  1. 普通区域分支：先建上下文再决定是否保留新建的 `AutoFilter`；`TryBuildContext` 失败时调用 `RollbackSyntheticAutoFilter(newFilterRange)`（无参 `AutoFilter()` 移除刚建立的筛选），保证不留下工作表变更。
  2. 调用顺序调整为 `newFilterRange.AutoFilter()` → `worksheet.AutoFilter` → `TryBuildContext(...)`；失败路径同时释放该 `AutoFilter` RCW。
  3. ListObject 分支：先读取并保存 `listObject.ShowAutoFilter` 原值，只有原值为 false 时才置 true 并记录 `changedShowAutoFilter`；`TryBuildContext` 失败时回滚为原值。用外层 `handedOff` 标志保证成功交接后不再由本方法释放上下文接管的 RCW。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0。代码级：`RollbackSyntheticAutoFilter` 被调用一次（失败路径）；`previousShowAutoFilter` / `changedShowAutoFilter` 回滚分支存在。

### E-08 `finally` 中恢复 Excel 状态可能掩盖原始异常
- **位置**：`BatchFilterService.cs:214-218,238-242`、`UnmergeAndFillService.cs:43-53`
- **方案**：每处恢复单独 try/catch 吞掉。
- **实际处置**：`BatchFilterService.Apply` 与 `BatchFilterService.Clear` 的 `finally` 中，`Cursor`、`StatusBar`、`Calculation`、`EnableEvents`、`ScreenUpdating` 五处恢复各自独立 `try { … } catch (Exception) { }`。`UnmergeAndFillService.UnmergeAndFillAll` 的 `finally` 中 `Calculation`、`EnableEvents`、`ScreenUpdating` 同样各自独立捕获（并取代原先的嵌套 `try/finally`，后者在 `EnableEvents` 恢复抛异常时仍会连带抛出）。`ComHelper.Release` 自身也吞掉释放异常，避免释放路径掩盖在途异常。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0。代码级：`BatchFilterService` 中 5 个恢复点、`UnmergeAndFillService` 中 3 个恢复点均为独立 `try/catch`。

### E-09 取消合并边界缺口
- **位置**：`UnmergeAndFillService.cs:164-165,176-180`
- **方案**：错误值保留语义（`Range.Value` 写回错误码或跳过并告警）；`Calculation = xlCalculationManual` 并在 `finally` 恢复；数组公式/受保护表给出明确提示。
- **实际处置**：
  1. **错误值**：新增 `IsErrorValue(object)`，按 `Excel.XlCVError` 枚举逐个比对 `Value2` 返回的负整数（兼容 `int`/`short`/`long`）；命中时取消合并后**直接返回**，不把错误码写成数字，单元格保持错误值语义。
  2. **手动计算**：`UnmergeAndFillAll` 设置 `Application.Calculation = xlCalculationManual`，并在 `finally` 中独立 `try/catch` 恢复原值。
  3. **数组公式（CSE）**：新增 `IsArrayFormula(Excel.Range)`，当 `FormulaArray` 非空且与单格 `Formula` 不同时判定为数组公式，抛出明确提示：“该合并区域包含数组公式（CSE），Excel 无法把数组公式填充到已取消合并的多个单元格。请先手动取消数组公式后重试。”
  4. **受保护工作表**：`mergeArea.UnMerge()` 与填充写入各自 `catch (COMException)` 并转换为可操作提示（“工作表可能受保护。请先取消工作表保护后重试。”）。
  5. **失败进度可见**：单个区域失败时用 `InvalidOperationException` 包装为“处理第 i/n 个合并区域时失败（已完成的区域不会自动回滚）：<原因>”，把原异常作为 `InnerException` 保留；这与重构前 `MainRibbon` 的“操作失败：…”提示链自然衔接。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0。代码级：`IsErrorValue`、`IsArrayFormula`、`xlCalculationManual` 设置与恢复、三处 `catch (COMException)` 提示均存在。

### E-10 README 构建命令缺证书指纹（实测失败）
- **位置**：`README.md:85-86`、`eWorkhelper.csproj:33-34`
- **证据**：实测 `error MSB4044: 未给任务"ManageCertificateStore"的必需参数"CertificateThumbprint"赋值`
- **方案**：README 补 `/p:ManifestCertificateThumbprint=<thumbprint>` 或说明 `IWORKHELPER_MANIFEST_CERT_THUMBPRINT`；并考虑 `SignManifests` 条件化。
- **实际处置**：
  1. `README.md`「编译」章节重写为三种等价写法：直接传 `/p:SignManifests=true /p:ManifestCertificateThumbprint=<thumbprint>`；设置 `IWORKHELPER_MANIFEST_CERT_THUMBPRINT` 环境变量（此时 `SignManifests` 自动为 true）；以及任务给定的完整验证命令。并显式写出缺指纹时的 `MSB4044` 报错文本，说明这是 VSTO 工具链固定要求而非项目配置错误。
  2. `eWorkhelper.csproj`：`SignManifests` 改为条件化——未显式指定时，有指纹则为 `true`，否则为 `false`；`ManifestCertificateThumbprint` 仍可从 `IWORKHELPER_MANIFEST_CERT_THUMBPRINT` 回落。
  3. `docs/RELEASE.md`、`docs/Development.md` 的构建命令同步改为 `/t:Rebuild` + 指纹参数（原 `RELEASE.md` 的 `/t:Restore,Rebuild` 对无包引用的 VSTO 工程没有意义），并补上环境变量写法。
- **状态**：已验证
- **验证**：
  - 带 `/p:SignManifests=true /p:ManifestCertificateThumbprint=EC8D…D594`：退出码 0，产出 `bin\Release\eWorkhelper.dll` + `.manifest` + `.vsto`。
  - 仅设 `$env:IWORKHELPER_MANIFEST_CERT_THUMBPRINT='EC8D…D594'`（不带 `SignManifests`、不带 `ManifestCertificateThumbprint`）：退出码 0。
  - **诚实说明**：VSTO 目标文件中的 `ManageCertificateStore` 任务在 `VisualStudioForApplicationsBuild` 内**无条件执行**，因此“完全不提供指纹”的构建仍会在该任务上失败（MSB4044）。已检查 `Microsoft.VisualStudio.Tools.Office.targets:150-168` 确认该任务没有条件属性，无法用 MSBuild 属性安全跳过。因此 E-10 的可行修复只能是“文档写清楚 + `SignManifests` 条件化避免误产出未签名包”，无法做到“无证书也能编译”。未执行“无指纹构建”的实验，因为它必然失败且属于工具链限制而非本次改动引入。

### E-11 Ribbon `OfficeId` 缺失 + 文案与文档不一致
- **位置**：`MainRibbon.Designer.cs:39-43,71`；文档 `README.md:13,49,52`、`docs/FEATURES.md:5`
- **方案**：自定义选项卡应使用 `RibbonControlIdType.Custom`（或补 `OfficeId`）；统一文档与界面文案（"工作助手"/"取消合并"）。
- **实际处置**：
  1. `MainRibbon.Designer.cs`：`tabEWorkHelper.ControlId.ControlIdType` 由 `RibbonControlIdType.Office` 改为 `RibbonControlIdType.Custom`，不再出现“声明为 Office 类型却不给 `OfficeId`”的不确定渲染状态（选项卡标签保持“工作助手”）。
  2. 文档全部对齐真实界面文案：`README.md` 改为“自定义‘工作助手’选项卡”“点击‘批量过滤’”“点击 Ribbon 中的‘取消合并’按钮”；`docs/FEATURES.md`、`RELEASE_NOTES.md` 同步为自定义“工作助手”选项卡 / “取消合并”。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0（`RibbonControlIdType.Custom` 属 `Microsoft.Office.Tools.Ribbon` 公开枚举，编译通过）。`grep RibbonControlIdType` 仅剩 `Custom` 一处；`grep "eWorkHelper\` 选项卡"` 在 README/FEATURES 中已无残留。真机渲染（自定义选项卡出现在 Ribbon 末端）列入 `docs/TESTING.md` 第 8 节。

### E-12 通配符与 AutoFilter 容量/长度上限未处理
- **位置**：`BatchFilterService.cs:466-480,494-542`
- **方案**：转义 `~ * ?`；校验单条件长度；匹配值超容量时给明确提示而非原始 COM 错误。
- **实际处置**：
  1. `EscapeWildcardCharacters(string)`：把 `~`、`*`、`?` 各自前置一个 `~` 再交给 `BuildFilterCriteria`，使字面量条件不被当作通配模式；空值仍映射为 `"="`。
  2. `ValidateConditions(IList<string>)`（`internal static`，窗体在调用 `Apply` 前先行校验）：单个条件超过 `MaxConditionLength = 8192` 字符时返回明确中文提示（含第几个条件与实测长度），不把超长文本交给 Excel。
  3. 容量：匹配到的唯一值超过 `MaxFilterValueCount = 10000` 时，`Apply` 不调用 `AutoFilter`，返回 `Applied = false` 与提示“匹配到 N 个唯一值，超过 Excel 筛选值列表上限（10000 项），未修改工作表。请增加条件以缩小范围。”，避免原始 COM 错误。
  4. 0 匹配不再用占位 GUID 隐藏整张表（见 E-06），因此删除了已无调用方的 `CreateMissingValue`。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0；0 warning / 0 error（确认删除 `CreateMissingValue` 后无“未使用成员”类问题）。`grep` 确认 `EscapeWildcardCharacters`、`MaxConditionLength`、`MaxFilterValueCount` 均被引用；两处校验分支存在。

## 低

### E-13 条件去重大小写敏感与匹配语义不一致
- **位置**：`BatchFilterService.cs:120` vs `:179-180,504`
- **方案**：去重改 `OrdinalIgnoreCase`。
- **实际处置**：`NormalizeConditions` 的 `HashSet<string>` 比较器由 `StringComparer.Ordinal` 改为 `StringComparer.OrdinalIgnoreCase`，与 `IsMatch` 的 `OrdinalIgnoreCase` 匹配语义一致（`abc` 与 `ABC` 现在被视为同一条件）。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0。`grep -n "StringComparer" BatchFilterService.cs` 显示去重、唯一值集合、签名去重处统一为 `OrdinalIgnoreCase`，仅 `AppliedFilterState.Matches` 的筛选签名比较保留 `StringComparison.Ordinal`（签名本身由 invariant 文本构成，应保持精确比较）。

### E-14 工作表校验只比名字
- **位置**：`BatchFilterService.cs:279`
- **方案**：同时校验所属工作簿。
- **实际处置**：新增 `IsSameWorksheet(Excel.Worksheet candidate, Excel.Worksheet expected)`：先比 `Name`（`Ordinal`），再取两者的 `Parent as Excel.Workbook` 并比较 `FullName`（`OrdinalIgnoreCase`），`COMException` 时按不匹配返回，两个 `Workbook` RCW 均在 `finally` 中释放。`TryCreateFilterRangeFromSelectedHeader` 通过 `selectedRange.Worksheet as Excel.Worksheet` 取到候选工作表（该 RCW 也释放），再调用该比较。原先“只比 `selectedRange.Worksheet.Name != worksheet.Name`”的写法（既漏比工作簿，又产生未释放的中间 RCW）已删除。
- **状态**：已验证
- **验证**：命令同 E-01 → 退出码 0。`grep -n "IsSameWorksheet"` 显示定义 1 处、调用 1 处；旧的单行 `Name` 比较已不存在。

### E-15 `OrdinalIgnoreCase` 并非完全大小写不敏感
- **位置**：`BatchFilterService.cs:504,517,529`；文档 `README.md:15`、`FEATURES.md:7`
- **方案**：文档补一句限定说明（土耳其语 i 等）。
- **实际处置**：`README.md`「已知限制」新增一条：“‘不区分大小写’通过 .NET `OrdinalIgnoreCase` 实现，并非完整的 Unicode 大小写折叠：土耳其语/阿塞拜疆语的 `i`/`İ` 等区域特有大小写对不会被视作同一个字符。” `docs/FEATURES.md`「已知限制」与 `RELEASE_NOTES.md`「Known limitations」同步补充同义说明。代码行为未改变（选择文档修复而非改成 `CurrentCultureIgnoreCase`，后者会让匹配结果随用户区域设置漂移，反而更难预测）。
- **状态**：已验证
- **验证**：文档改动，`grep -n "OrdinalIgnoreCase" README.md docs/FEATURES.md RELEASE_NOTES.md` 均命中新说明；构建不受影响（E-01 命令退出码 0）。

### E-16 状态作用域文档漂移
- **位置**：`MainRibbon.cs:10` vs `README.md:97`、`RELEASE_NOTES.md:48`、`FEATURES.md:30`
- **方案**：文档改为“每个 Excel 窗口（Ribbon 实例）”，或将服务提升为进程级单例。
- **实际处置**：采用文档修复（未提升为进程级单例，避免行为变更与跨窗口状态耦合）。`README.md` 的已知限制改为“保存在加载项进程内存中，且**每个 Excel 窗口（每个 Ribbon 实例）各有一份**：同一进程内的其他窗口看不到该状态，加载项卸载后状态也不会保留”；`docs/FEATURES.md`、`RELEASE_NOTES.md`、`docs/ARCHITECTURE.md`（新增“状态作用域”条目）同步说明；`docs/ARCHITECTURE.md` 同时记录了“服务内只保存身份信息，不缓存 Workbook/Worksheet RCW”（与 E-03 一致）。
- **状态**：已验证
- **验证**：文档改动；`grep -n "每个 Excel 窗口" README.md docs/FEATURES.md docs/ARCHITECTURE.md RELEASE_NOTES.md` 均命中。构建不受影响。

### E-17 死代码与冗余配置清理
- 原列 7 项，逐项处置：
  1. **`BatchFilterForm.cs:208-218` `InitializeComponent()` 从未调用** → 已确认全仓仅此一处定义、无任何调用点（`grep -n "InitializeComponent"` 只命中 `MainRibbon.Designer.cs` 与 `ThisAddIn.Designer.cs` 的同名方法）；已删除整个方法。构造函数负责构建全部控件并设置 `ClientSize = new Size(520, 410)`，删除后布局与 `ClientSize` **完全不变**（原方法设置的 `ClientSize(284,261)` 本来就因未被调用而无效）。
  2. **`Properties\Resources` 的 `Filter` 与 `Resources\Filter.png` 未被引用** → 已确认：`grep` 全仓 `Filter.png` / `Resources.Filter` 仅命中 `Resources.resx`、`Resources.Designer.cs`、`eWorkhelper.csproj`（即定义方自身），Ribbon 的实际按钮图片内嵌在 `MainRibbon.resx` 的 `btnBatchFilter.Image`（base64，与 `Filter.png` 无关）；已删除 `Resources.resx` 的 `Filter` 数据项、`Resources.Designer.cs` 的 `Filter` 属性、`eWorkhelper.csproj` 的 `<None Include="Resources\Filter.png" />`，并删除 `Resources\Filter.png`。
  3. **`eWorkhelper.csproj` `.NET Framework 3.5 SP1` bootstrapper（`Install=false`）** → 已删除该 `BootstrapperPackage` 项。
  4. **`Properties\Settings.settings` 为空，无任何读取方** → 已确认 `Settings.Designer.cs` 只有 `Default` 单例、全仓无 `Settings.Default` 调用点；删除 `Settings.settings` 与 csproj 中对应的 `<None>` 项，保留 `Settings.Designer.cs` 编译项。
  5. **`AutoIncrementApplicationRevision=true`** → 改为 `false`，发布版本可复现。
  6. **Debug `WarningLevel=5`（C# 合法范围 0–4）** → 改为 `4`，与 Release 一致。
  7. **`VisualStudioVersion` 空值时 `VSToolsPath` 回落 v10.0** → 删除 `VisualStudioVersion` 的 `10.0` 回落，并把 `VSToolsPath` 的条件收紧为 `'$(VisualStudioVersion)' != ''`，不再可能拼出 `…\Microsoft\VisualStudio\v10.0`。
- **状态**：已验证
- **验证**：
  - 命令同 E-01 → 退出码 0，产出 `eWorkhelper.dll` / `.manifest` / `.vsto`；Release 与 Debug 均 0 warning / 0 error（改动 6、7 直接影响 Release/Debug 编译参数，两配置均已实测）。
  - `glob` 结果确认仓库内已无 `Resources\Filter.png`、`Properties\Settings.settings`；`Get-ChildItem docs` 确认 4 个空目录已删除。
  - `grep -c "InitializeComponent" BatchFilterForm.cs` = 0；`grep -c "Settings.Default" **/*.cs` = 0。

### E-18 文档结构整理
- 原列 4 项，逐项处置：
  1. **根 `RELEASE_NOTES.md` 与 `docs/CHANGELOG.md` 两条并行发布历史 → 合并为单一权威**：`docs/CHANGELOG.md` 声明为**唯一权威发布历史**；`RELEASE_NOTES.md` 顶部加入醒目指引，只保留当前版本（v1.2.0）的对外摘要并声明“请勿在本文件中单独新增版本条目”。`README.md`「文档入口」与 `docs/Development.md`「文档导航」均写明 CHANGELOG 为唯一权威。
  2. **`docs/CHANGELOG.md` 补 `v1.1.260822.6` 条目、加日期**：已新增 `## v1.1.260822.6 — 日期 2026-08-22（按版本号 260822 段推断）` 条目，并给 `v1.2.0` 标注日期 `2026-08-22`。**诚实说明**：仓库内没有该版本的历史发布记录，日期由版本号中的 `260822` 段推断，条目正文属基于现有代码与文档的重建，并非从原始发布记录抄录；已在标题中显式标注“推断”。
  3. **删除空目录 `docs/Architecture/`、`docs/Features/`、`docs/Release/`、`docs/Testing/`**：删除前逐个确认 `Get-ChildItem -Force` 计数为 0（四个均为 0 项），随后删除；`git ls-files docs` 也确认这四个目录从未入库。
  4. **补 `docs/TESTING.md`**：新增完整人工回归清单，共 9 节 —— 环境与前置、构建与签名、批量过滤范围识别、匹配语义（含日期/百分比/货币/长文本/通配符/大小写/空值）、结果与清除、进度与取消、取消合并并填充（含错误值/数组公式/受保护表/状态恢复）、COM 生命周期（进程能否干净退出、句柄是否持续增长）、Ribbon 与关于，末尾附回归记录表。`docs/RELEASE.md`、`docs/Development.md`、`README.md`、`docs/CHANGELOG.md` 均已链接到它。
- **状态**：已验证
- **验证**：`Get-ChildItem docs` 列出 `ARCHITECTURE.md`、`CHANGELOG.md`、`Development.md`、`FEATURES.md`、`RELEASE.md`、`REVIEW_TRACKING.md`、`TESTING.md`；四个空目录已不存在。`grep -n "TESTING.md"` 命中 README、RELEASE、Development、CHANGELOG 的链接。文档改动不影响构建（E-01 命令退出码 0）。

---

## 变更记录

| 日期 | 内容 |
|---|---|
| 2026-09-13 | 建立跟踪文档，录入全部条目（E-01…E-18） |
| 2026-09-13 | E-01：新增 `ComHelper.cs`；`BatchFilterService.cs`、`UnmergeAndFillService.cs`、`BatchFilterForm.cs`、`MainRibbon.cs` 全量补齐 COM 释放（双点链拆局部变量 + `finally` 逆序释放）。Release/Debug 构建 0 warning / 0 error。 |
| 2026-09-13 | E-03：`AppliedFilterState` 改存工作簿全名/工作表名/行列号身份信息；`BatchFilterContext` 实现 `IDisposable`；`BatchFilterForm.Dispose(bool)` 释放上下文，`MainRibbon` 在 `finally` 中 `form.Dispose()`。 |
| 2026-09-13 | E-02 + E-06：匹配与可见值读取统一改用 `Range.Text` 显示文本（`#####` 溢出回退 + 百分比/区域格式辅助候选）；`LoadInitialState` 与 `Apply` 语义对齐；新增 `VisibleCount` 并按需区分“条件匹配行数”与“实际可见行数”；0 匹配不再隐藏整张表。 |
| 2026-09-13 | E-04：`LoadInitialState` 与 `btnBatchFilter_Click` 补 try/catch；`GetTargetFilter` 守护 `Filters` 访问。 |
| 2026-09-13 | E-05：加入 `Application.Cursor(xlWait)`、`StatusBar` 进度、`DoEvents` 泵消息、协作式取消（Esc / `CancellationTokenSource`）、`busy` 重入防护与 `OnFormClosing` 守护；全部恢复动作放在 `finally`。 |
| 2026-09-13 | E-07 + E-08：新建 AutoFilter 失败时回滚；ListObject 的 `ShowAutoFilter` 原值保存与回滚；所有 Excel 状态恢复改为各自独立 `try/catch`。 |
| 2026-09-13 | E-09：`UnmergeAndFillService` 识别并跳过错误值、设置/恢复 `xlCalculationManual`、对数组公式（CSE）与受保护表给出明确提示，并在失败提示中报告已完成区域数。 |
| 2026-09-13 | E-12 + E-13 + E-14：新增通配符转义 `EscapeWildcardCharacters`、条件长度（8192）与值列表容量（10000）校验；条件去重改 `OrdinalIgnoreCase`；工作表校验改为工作簿 + 工作表。 |
| 2026-09-13 | E-11：Ribbon 选项卡改用 `RibbonControlIdType.Custom`；README/FEATURES/RELEASE_NOTES 文案与界面统一为“工作助手”“取消合并”。 |
| 2026-09-13 | E-10 + E-17：`SignManifests` 条件化；README/RELEASE/Development 构建命令补齐指纹与 `IWORKHELPER_MANIFEST_CERT_THUMBPRINT` 说明；删除未调用的 `InitializeComponent()` 与未使用的 `Filter` 资源、`.NET 3.5 SP1` bootstrapper、空 `Settings.settings`；`AutoIncrementApplicationRevision=false`；Debug `WarningLevel=4`；移除 `VisualStudioVersion` 的 v10.0 回落。 |
| 2026-09-13 | E-15 + E-16 + E-18：补充 `OrdinalIgnoreCase` 与土耳其语 i 的限定说明；状态作用域改为“每个 Excel 窗口（每个 Ribbon 实例）”；`docs/CHANGELOG.md` 定为唯一权威发布历史并补 `v1.1.260822.6`；`RELEASE_NOTES.md` 改为指向性摘要；删除 4 个空文档目录；新增 `docs/TESTING.md` 人工回归清单。 |
| 2026-09-13 | 最终验证：Release 与 Debug 均 `Rebuild` 退出码 0、0 warning / 0 error；环境变量路径亦退出码 0。 |

---

## 未完成 / 需人工确认的事项

| 事项 | 原因 |
|---|---|
| E-01 / E-02 / E-05 / E-06 / E-07 / E-09 / E-11 的**运行期**行为 | 本会话无法启动带 VSTO 加载项的 Excel 做真机交互（无自动化测试工程、无法在无用户交互下驱动 Excel UI）。已按 `docs/TESTING.md` 列出逐条人工验证步骤，需由维护者在真机 Excel 上执行。 |
| “完全不提供证书指纹也能编译” | `Microsoft.VisualStudio.Tools.Office.targets:150-168` 的 `ManageCertificateStore` 任务无条件执行且必需 `CertificateThumbprint`，MSBuild 层无法安全跳过。已改为文档写清 + `SignManifests` 条件化。 |
| `v1.1.260822.6` 条目正文 | 仓库内无该版本原始发布记录；条目按版本号 `260822` 推断日期、按现有代码与文档重建，已在 CHANGELOG 标题中标注“推断”。 |
