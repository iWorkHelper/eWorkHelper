# 取消合并并填充

## 功能名称与目标

Excel Ribbon 的 `eWorkHelper` Tab 在“数据工具”Group 中提供第二个正式功能“取消合并并填充”。用户执行一次操作，即可取消当前 Selection 涉及的全部独立合并区域，并把原合并单元格内容填充到原区域的每个单元格。普通单元格不得修改。

## 操作流程

1. 用户在当前工作表选择一个连续范围或通过 Ctrl 选择多个 Areas。
2. 点击“取消合并并填充”。
3. 服务遍历所有 Areas，只发现与选区相交的 MergeArea；即使只选择合并区域的一部分，也处理完整 MergeArea。
4. 每个 MergeArea 在本次操作中去重，取消合并前读取左上角内容。
5. 调用 Excel 原生 `Range.UnMerge()`，再把原内容写入原 MergeArea 全部单元格。
6. 完成后显示处理区域数量；没有合并单元格时仅提示，不修改工作表。

## Excel Object Model 实现方案

- Ribbon 从 `Globals.ThisAddIn.Application` 获取宿主 Application，服务校验 `Application.Selection` 是否为 `Excel.Range`。
- 遍历 `Selection.Areas`。每个 Area 先读取 `MergeCells`：明确为 `false` 的块立即跳过；混合或包含合并的块按行列较长方向递归拆分，直到定位到属于 MergeArea 的单元格。这样稀疏大范围无需逐个检查所有普通单元格。
- 对命中的单元格获取完整 `MergeArea`，使用工作表运行时身份和 `MergeArea.Address` 的外部绝对地址组成唯一键。Ctrl 多选、重叠 Areas 或同一合并区域中的多个单元格都不会导致重复处理。
- 在 `UnMerge()` 前从 MergeArea 左上角读取 `HasFormula`、`FormulaR1C1` 或 `Value2`，并保存原范围对象。
- 普通内容在取消合并后通过整块 `Range.Value2 = originalValue` 一次性填充。空内容保留为真正空值，不写入字符串 `"null"`。
- 公式通过整块 `Range.FormulaR1C1 = originalFormulaR1C1` 写回。标量 R1C1 公式在每个目标单元格保留相同的相对引用模式，等价于将原公式复制到区域各单元格；绝对引用保持绝对，且不会把公式结果转成固定值。此策略可能使相对引用随目标单元格位置变化，这是 Excel 复制公式的原生语义。
- 不创建合并区域、辅助行列或工作表，不剪切粘贴、不隐藏单元格，也不主动设置任何格式。

## 性能与状态安全

- 通过范围级 `MergeCells` 判断和递归剪枝减少跨 COM 边界调用；发现 MergeArea 后立即去重，不重复读取、取消合并或写入。
- 操作期间只临时关闭 `ScreenUpdating` 和 `EnableEvents`，并在 `finally` 中恢复用户原始状态。Calculation 不修改。
- 单个 MergeArea 的内容读取发生在 UnMerge 前，写入使用整块 Range 赋值，不逐单元格写入。

## 边界条件与异常处理

- 没有工作簿、Selection 不是 Range 或当前选择不可处理时，不修改工作表并返回简短错误信息。
- Selection 没有合并单元格时返回处理数量 0，不抛异常。
- Selection 混合普通单元格与合并单元格时只处理后者。
- Selection 只触及 MergeArea 一部分时仍处理完整区域。
- Excel Selection 正常只属于当前工作表，本功能不扩展跨工作表选择。
- 工作表保护、只读限制、无权限或 COM 操作失败时，服务恢复 Application 临时状态后让异常返回 Ribbon 层；Ribbon 使用现有 MessageBox 规范显示简短错误信息，不引入新日志框架。当前项目没有日志机制，因此不伪造或吞掉异常。
- Excel 原生 UnMerge 本身不提供可由 VSTO 合并进 Excel 撤销栈的事务能力；若多个区域处理中后续区域失败，已完成的区域不能自动回滚。Application 状态一定恢复，错误会明确反馈。

## 验收标准

- 单个纵向、横向、矩形 MergeArea 取消合并后，每个单元格都包含原内容。
- 一次选择多个连续合并区域、大范围中的多个合并区域或 Ctrl 多选 Areas 时，全部独立 MergeArea 一次处理完成。
- 同一 MergeArea 无论被多少选区单元格命中都只执行一次。
- 文本、数字、日期、布尔值、空值通过 `Value2` 正常处理；公式保持为公式并遵循 R1C1 复制语义。
- 普通单元格、工作表结构及格式属性不被功能主动修改。
- 无合并区域不报错；异常路径恢复 ScreenUpdating 和 EnableEvents 并提示用户。
- Debug 与 Release 全量构建均为 0 Error；本次引入的 Warning 为 0。

Excel 官方说明确认，对多单元格 Range 设置 `FormulaR1C1` 会使用该公式填充所有单元格，详见 [Range.FormulaR1C1 property](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formular1c1)。

## 手工验证步骤

1. 新建工作簿，分别准备纵向、横向、矩形、多个分散合并区域以及混合普通单元格的大范围。
2. 对文本、数字、日期、布尔值、空值和包含相对/绝对引用的公式分别执行功能，核对每个目标单元格的值或公式。
3. 分别测试连续选择、Ctrl 多选、只选 MergeArea 内一个非左上角单元格和无合并区域。
4. 对 A1:Z5000 中少量合并区域执行，观察性能并确认普通单元格不变。
5. 在受保护工作表触发失败，确认提示出现且 ScreenUpdating、EnableEvents 恢复。
6. 比对操作前后格式、行高和列宽；再完成 Excel F5 启动、Add-in 加载、Ribbon 布局及两个正式功能按钮回归。

## 当前验证结果

- Debug / Any CPU Rebuild：通过，0 Error，0 Warning。
- Release / Any CPU Rebuild：通过，0 Error，0 Warning。
- Debug、Release ProductVersion：正式发布构建完成后核验为 `1.1.260822.6`。
- Ribbon 与源码静态检查：通过；无 Test/Demo/Debug 入口，事件绑定、单一 `UnMerge()` 调用点、Value2/FormulaR1C1 分支和 finally 状态恢复均存在。
- 自动化测试：项目没有独立测试项目，未新增依赖 Excel COM 的脆弱伪单元测试。
- Excel F5/GUI 和真实工作簿场景：本次未执行，以上 17 类 Excel 行为及格式、异常路径需按回归清单人工验证。

## 实现位置与已知限制

- Ribbon 入口：`MainRibbon.Designer.cs`、`MainRibbon.cs`。
- 核心服务：`UnmergeAndFillService.cs`。
- 项目注册：`eWorkhelper.csproj`。
- Excel 不向 VSTO 暴露可将多次 UnMerge 和写入组合为单一事务撤销单元；后续区域失败时，之前已完成的区域不能自动回滚。
