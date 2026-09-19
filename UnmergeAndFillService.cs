using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;

namespace eWorkhelper
{
    internal sealed class UnmergeAndFillService
    {
        internal int Execute(Excel.Application application)
        {
            if (application == null)
            {
                throw new InvalidOperationException("无法获取当前 Excel 应用程序。");
            }

            // application 是 VSTO 宿主项，归运行时所有，绝不释放。
            Excel.Range selection = null;
            try
            {
                selection = application.Selection as Excel.Range;
                if (selection == null)
                {
                    throw new InvalidOperationException("请先选择要处理的单元格区域。");
                }

                IList<Excel.Range> mergeAreas = FindMergeAreas(selection);
                try
                {
                    if (mergeAreas.Count == 0)
                    {
                        return 0;
                    }

                    return UnmergeAndFillAll(application, mergeAreas);
                }
                finally
                {
                    ReleaseAreas(mergeAreas);
                }
            }
            finally
            {
                ComHelper.Release(selection);
            }
        }

        private static int UnmergeAndFillAll(Excel.Application application, IList<Excel.Range> mergeAreas)
        {
            bool originalScreenUpdating = true;
            bool haveScreenUpdating = false;
            bool originalEnableEvents = true;
            bool haveEnableEvents = false;
            Excel.XlCalculation originalCalculation = Excel.XlCalculation.xlCalculationAutomatic;
            bool haveCalculation = false;

            try
            {
                // E-08：每个状态读取/恢复都单独保护，避免失败替换掉在途异常。
                try
                {
                    originalScreenUpdating = application.ScreenUpdating;
                    haveScreenUpdating = true;
                    application.ScreenUpdating = false;
                }
                catch (COMException)
                {
                    haveScreenUpdating = false;
                }

                try
                {
                    originalEnableEvents = application.EnableEvents;
                    haveEnableEvents = true;
                    application.EnableEvents = false;
                }
                catch (COMException)
                {
                    haveEnableEvents = false;
                }

                // E-09：逐个区域写公式/值会触发整表重算，改为手动计算并在 finally 恢复。
                try
                {
                    originalCalculation = application.Calculation;
                    haveCalculation = true;
                    application.Calculation = Excel.XlCalculation.xlCalculationManual;
                }
                catch (COMException)
                {
                    haveCalculation = false;
                }

                int processed = 0;
                try
                {
                    foreach (Excel.Range mergeArea in mergeAreas)
                    {
                        UnmergeAndFill(mergeArea);
                        processed++;
                    }
                }
                catch (Exception ex)
                {
                    // 该操作不是 Excel 事务，已完成区域无法自动回滚；把进度如实告知用户。
                    throw new InvalidOperationException(
                        string.Format(
                            CultureInfo.CurrentCulture,
                            "处理第 {0}/{1} 个合并区域时失败（已完成的区域不会自动回滚）：{2}",
                            processed + 1,
                            mergeAreas.Count,
                            ex.Message),
                        ex);
                }

                return processed;
            }
            finally
            {
                if (haveCalculation)
                {
                    try
                    {
                        application.Calculation = originalCalculation;
                    }
                    catch (Exception)
                    {
                    }
                }

                if (haveEnableEvents)
                {
                    try
                    {
                        application.EnableEvents = originalEnableEvents;
                    }
                    catch (Exception)
                    {
                    }
                }

                if (haveScreenUpdating)
                {
                    try
                    {
                        application.ScreenUpdating = originalScreenUpdating;
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static IList<Excel.Range> FindMergeAreas(Excel.Range selection)
        {
            List<Excel.Range> mergeAreas = new List<Excel.Range>();
            HashSet<string> processedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            int areaCount = CountAreas(selection);
            for (int areaIndex = 1; areaIndex <= areaCount; areaIndex++)
            {
                Excel.Range area = GetArea(selection, areaIndex);
                if (area == null)
                {
                    continue;
                }

                try
                {
                    FindMergeAreasInRange(area, processedKeys, mergeAreas);
                }
                finally
                {
                    ComHelper.Release(area);
                }
            }

            return mergeAreas;
        }

        private static void FindMergeAreasInRange(
            Excel.Range range,
            ISet<string> processedKeys,
            IList<Excel.Range> mergeAreas)
        {
            object mergeState = range.MergeCells;
            if (mergeState is bool && !(bool)mergeState)
            {
                return;
            }

            int rowCount;
            int columnCount;
            int firstRow;
            int firstColumn;
            Excel.Range rangeRows = null;
            Excel.Range rangeColumns = null;
            Excel.Range firstCell = null;
            Excel.Range firstMergeArea = null;
            bool firstMergeAreaOwned = false;
            try
            {
                rangeRows = range.Rows;
                rowCount = rangeRows.Count;
                rangeColumns = range.Columns;
                columnCount = rangeColumns.Count;
                firstRow = range.Row;
                firstColumn = range.Column;

                firstCell = range.Cells[1, 1] as Excel.Range;
                if (firstCell == null)
                {
                    return;
                }

                bool firstCellMerged = false;
                try
                {
                    firstCellMerged = (bool)firstCell.MergeCells;
                }
                catch (COMException)
                {
                    firstCellMerged = false;
                }

                if (firstCellMerged)
                {
                    firstMergeArea = firstCell.MergeArea;
                    if (firstMergeArea != null && ContainsRange(firstMergeArea, firstRow, firstColumn, rowCount, columnCount))
                    {
                        firstMergeAreaOwned = AddMergeArea(firstMergeArea, processedKeys, mergeAreas);
                        return;
                    }
                }
                else
                {
                    firstMergeArea = null;
                }

                if (rowCount == 1 && columnCount == 1)
                {
                    if ((bool)range.MergeCells)
                    {
                        Excel.Range mergeArea = range.MergeArea;
                        if (mergeArea != null)
                        {
                            bool owned = AddMergeArea(mergeArea, processedKeys, mergeAreas);
                            if (!owned)
                            {
                                ComHelper.Release(mergeArea);
                            }
                        }
                    }

                    return;
                }

                if (rowCount >= columnCount && rowCount > 1)
                {
                    int firstRowCount = rowCount / 2;
                    Excel.Range firstHalf = range.Resize[firstRowCount, columnCount];
                    Excel.Range offsetRange = range.Offset[firstRowCount, 0];
                    Excel.Range secondHalf = null;
                    try
                    {
                        secondHalf = offsetRange.Resize[rowCount - firstRowCount, columnCount];
                        FindMergeAreasInRange(firstHalf, processedKeys, mergeAreas);
                        FindMergeAreasInRange(secondHalf, processedKeys, mergeAreas);
                    }
                    finally
                    {
                        ComHelper.Release(secondHalf);
                        ComHelper.Release(offsetRange);
                        ComHelper.Release(firstHalf);
                    }

                    return;
                }

                int firstColumnCount = columnCount / 2;
                Excel.Range leftHalf = range.Resize[rowCount, firstColumnCount];
                Excel.Range columnOffset = range.Offset[0, firstColumnCount];
                Excel.Range rightHalf = null;
                try
                {
                    rightHalf = columnOffset.Resize[rowCount, columnCount - firstColumnCount];
                    FindMergeAreasInRange(leftHalf, processedKeys, mergeAreas);
                    FindMergeAreasInRange(rightHalf, processedKeys, mergeAreas);
                }
                finally
                {
                    ComHelper.Release(rightHalf);
                    ComHelper.Release(columnOffset);
                    ComHelper.Release(leftHalf);
                }
            }
            finally
            {
                // 只有未被 mergeAreas 接管的合并区域才在这里释放。
                if (!firstMergeAreaOwned)
                {
                    ComHelper.Release(firstMergeArea);
                }

                ComHelper.Release(firstCell);
                ComHelper.Release(rangeColumns);
                ComHelper.Release(rangeRows);
            }
        }

        /// <summary>
        /// 记录一个合并区域。返回 true 表示 <paramref name="mergeArea"/> 已被 <paramref name="mergeAreas"/> 接管，
        /// 调用方不得再释放它。
        /// </summary>
        private static bool AddMergeArea(
            Excel.Range mergeArea,
            ISet<string> processedKeys,
            IList<Excel.Range> mergeAreas)
        {
            string key = mergeArea.get_Address(
                true,
                true,
                Excel.XlReferenceStyle.xlA1,
                true,
                Type.Missing);

            if (processedKeys.Add(key))
            {
                mergeAreas.Add(mergeArea);
                return true;
            }

            return false;
        }

        private static bool ContainsRange(
            Excel.Range container,
            int candidateFirstRow,
            int candidateFirstColumn,
            int candidateRowCount,
            int candidateColumnCount)
        {
            int containerLastRow;
            int containerLastColumn;
            Excel.Range containerRows = null;
            Excel.Range containerColumns = null;
            try
            {
                containerRows = container.Rows;
                containerColumns = container.Columns;
                containerLastRow = container.Row + containerRows.Count - 1;
                containerLastColumn = container.Column + containerColumns.Count - 1;
            }
            finally
            {
                ComHelper.Release(containerColumns);
                ComHelper.Release(containerRows);
            }

            int candidateLastRow = candidateFirstRow + candidateRowCount - 1;
            int candidateLastColumn = candidateFirstColumn + candidateColumnCount - 1;

            return candidateFirstRow >= container.Row
                && candidateFirstColumn >= container.Column
                && candidateLastRow <= containerLastRow
                && candidateLastColumn <= containerLastColumn;
        }

        private static void UnmergeAndFill(Excel.Range mergeArea)
        {
            Excel.Range topLeft = null;
            try
            {
                topLeft = mergeArea.Cells[1, 1] as Excel.Range;
                if (topLeft == null)
                {
                    throw new InvalidOperationException("无法读取合并区域的左上角单元格。");
                }

                if (IsArrayFormula(topLeft))
                {
                    throw new InvalidOperationException(
                        "该合并区域包含数组公式（CSE），Excel 无法把数组公式填充到已取消合并的多个单元格。请先手动取消数组公式后重试。");
                }

                bool hasFormula = HasFormula(topLeft);
                object content = hasFormula ? topLeft.FormulaR1C1 : topLeft.Value2;
                bool isError = !hasFormula && IsErrorValue(content);

                try
                {
                    mergeArea.UnMerge();
                }
                catch (COMException ex)
                {
                    throw new InvalidOperationException(
                        "无法取消合并该区域，工作表可能受保护。请先取消工作表保护后重试。", ex);
                }

                if (content == null || isError)
                {
                    // E-09：错误值（#DIV/0! 等的 Value2 是负错误码）不写回，避免把错误码落成数字。
                    return;
                }

                try
                {
                    if (hasFormula)
                    {
                        mergeArea.FormulaR1C1 = content;
                    }
                    else
                    {
                        mergeArea.Value2 = content;
                    }
                }
                catch (COMException ex)
                {
                    throw new InvalidOperationException(
                        "填充取消合并后的单元格失败，工作表可能受保护或区域为只读。请先取消工作表保护后重试。", ex);
                }
            }
            finally
            {
                ComHelper.Release(topLeft);
            }
        }

        private static bool HasFormula(Excel.Range cell)
        {
            try
            {
                object hasFormula = cell.HasFormula;
                return hasFormula is bool && (bool)hasFormula;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static bool IsArrayFormula(Excel.Range cell)
        {
            try
            {
                object formula = cell.FormulaArray;
                string arrayFormula = formula as string;
                if (string.IsNullOrEmpty(arrayFormula))
                {
                    return false;
                }

                string singleFormula = null;
                try
                {
                    singleFormula = cell.Formula as string;
                }
                catch (COMException)
                {
                    singleFormula = null;
                }

                // 仅当存在独立的单格公式且与数组公式不同，才判定为 CSE 数组公式。
                return !string.IsNullOrEmpty(singleFormula)
                    && !string.Equals(singleFormula, arrayFormula, StringComparison.Ordinal);
            }
            catch (COMException)
            {
                return false;
            }
        }

        /// <summary>
        /// E-09：识别错误值。Excel 错误值的 <c>Value2</c> 是负的 <see cref="Excel.XlCVError"/> 枚举值。
        /// </summary>
        private static bool IsErrorValue(object value)
        {
            if (value == null)
            {
                return false;
            }

            int code;
            if (value is int)
            {
                code = (int)value;
            }
            else if (value is short)
            {
                code = (short)value;
            }
            else if (value is long)
            {
                code = (int)(long)value;
            }
            else
            {
                return false;
            }

            if (code >= 0)
            {
                return false;
            }

            foreach (Excel.XlCVError errorValue in Enum.GetValues(typeof(Excel.XlCVError)))
            {
                if ((int)errorValue == code)
                {
                    return true;
                }
            }

            return false;
        }

        private static int CountAreas(Excel.Range range)
        {
            Excel.Areas areas = null;
            try
            {
                areas = range.Areas;
                return areas == null ? 0 : areas.Count;
            }
            catch (COMException)
            {
                return 0;
            }
            finally
            {
                ComHelper.Release(areas);
            }
        }

        private static Excel.Range GetArea(Excel.Range range, int oneBasedIndex)
        {
            Excel.Areas areas = null;
            try
            {
                areas = range.Areas;
                return areas == null ? null : areas[oneBasedIndex] as Excel.Range;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                ComHelper.Release(areas);
            }
        }

        private static void ReleaseAreas(IList<Excel.Range> areas)
        {
            if (areas == null)
            {
                return;
            }

            for (int index = areas.Count - 1; index >= 0; index--)
            {
                ComHelper.Release(areas[index]);
            }
        }
    }
}
