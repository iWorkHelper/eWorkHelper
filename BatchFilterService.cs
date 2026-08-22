using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace eWorkhelper
{
    internal enum BatchFilterMatchMode
    {
        Equals,
        NotEquals,
        Contains,
        NotContains
    }

    internal sealed class BatchFilterContext
    {
        internal Excel.Application Application { get; set; }
        internal Excel.Workbook Workbook { get; set; }
        internal Excel.Worksheet Worksheet { get; set; }
        internal Excel.AutoFilter AutoFilter { get; set; }
        internal Excel.Range FilterRange { get; set; }
        internal Excel.Range DataRange { get; set; }
        internal string ColumnDisplayName { get; set; }
        internal int FieldIndex { get; set; }
        internal int TargetColumn { get; set; }
        internal int HeaderRow { get; set; }
        internal int DataRowCount { get; set; }
    }

    internal sealed class BatchFilterResult
    {
        internal int CheckedCount { get; set; }
        internal int MatchedCount { get; set; }
    }

    internal sealed class BatchFilterInitialState
    {
        internal IList<string> Conditions { get; set; }
        internal BatchFilterMatchMode MatchMode { get; set; }
        internal string StatusMessage { get; set; }
    }

    internal sealed class BatchFilterService
    {
        private AppliedFilterState currentState;

        internal bool TryCreateContext(Excel.Application application, out BatchFilterContext context, out string errorMessage)
        {
            context = null;
            errorMessage = null;
            try
            {
                if (application == null || application.Workbooks.Count == 0)
                {
                    errorMessage = "没有可用的 Excel 工作簿。";
                    return false;
                }

                Excel.Worksheet worksheet = application.ActiveSheet as Excel.Worksheet;
                Excel.Range activeCell = application.ActiveCell as Excel.Range;
                if (worksheet == null || activeCell == null)
                {
                    errorMessage = "没有有效的工作表或活动单元格。";
                    return false;
                }

                int targetColumn = activeCell.Column;
                Excel.ListObject listObject = activeCell.ListObject;
                if (listObject != null)
                {
                    return TryCreateListObjectContext(application, worksheet, listObject, targetColumn, out context, out errorMessage);
                }

                if (worksheet.AutoFilterMode)
                {
                    Excel.AutoFilter autoFilter = worksheet.AutoFilter;
                    Excel.Range existingRange = autoFilter == null ? null : autoFilter.Range;
                    if (existingRange == null || !ContainsColumn(existingRange, targetColumn))
                    {
                        errorMessage = "当前选择的列不在现有筛选区域内，请选择筛选区域中的列后重试。";
                        return false;
                    }

                    return TryBuildContext(application, worksheet, autoFilter, existingRange, targetColumn, null, out context, out errorMessage);
                }

                Excel.Range newFilterRange;
                bool cancelled;
                if (!TryCreateFilterRangeFromSelectedHeader(application, worksheet, targetColumn, out newFilterRange, out cancelled, out errorMessage))
                {
                    return false;
                }

                if (cancelled)
                {
                    return false;
                }

                newFilterRange.AutoFilter();
                return TryBuildContext(application, worksheet, worksheet.AutoFilter, newFilterRange, targetColumn, null, out context, out errorMessage);
            }
            catch (COMException)
            {
                errorMessage = "无法识别或建立 Excel 筛选区域，请确认工作表未受保护并重试。";
                return false;
            }
            catch (Exception ex)
            {
                errorMessage = "无法启动批量过滤：" + ex.Message;
                return false;
            }
        }

        internal IList<string> NormalizeConditions(string input)
        {
            List<string> conditions = new List<string>();
            HashSet<string> unique = new HashSet<string>(StringComparer.Ordinal);
            string[] lines = (input ?? string.Empty).Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            foreach (string line in lines)
            {
                string condition = line.Trim();
                if (condition.Length > 0 && unique.Add(condition))
                {
                    conditions.Add(condition);
                }
            }

            return conditions;
        }

        internal BatchFilterInitialState LoadInitialState(BatchFilterContext context)
        {
            Excel.Filter filter = GetTargetFilter(context);
            if (filter == null || !filter.On)
            {
                return new BatchFilterInitialState
                {
                    Conditions = new List<string>(),
                    MatchMode = BatchFilterMatchMode.Contains,
                    StatusMessage = "请输入过滤条件。"
                };
            }

            string filterSignature = TryGetFilterSignature(filter);
            if (currentState != null && currentState.Matches(context, filterSignature))
            {
                return new BatchFilterInitialState
                {
                    Conditions = currentState.GetConditions(),
                    MatchMode = currentState.MatchMode,
                    StatusMessage = "已恢复上次批量过滤条件。"
                };
            }

            IList<string> visibleValues = ReadVisibleUniqueValues(context.DataRange);
            return new BatchFilterInitialState
            {
                Conditions = visibleValues,
                MatchMode = BatchFilterMatchMode.Equals,
                StatusMessage = visibleValues.Count == 0
                    ? "当前筛选无可见数据。"
                    : string.Format("已加载当前筛选的 {0} 个可见唯一值。", visibleValues.Count)
            };
        }

        internal BatchFilterResult Apply(BatchFilterContext context, IList<string> conditions, BatchFilterMatchMode mode)
        {
            bool previousScreenUpdating = context.Application.ScreenUpdating;
            bool previousEnableEvents = context.Application.EnableEvents;
            try
            {
                context.Application.ScreenUpdating = false;
                context.Application.EnableEvents = false;

                object values = context.DataRange.Value2;
                HashSet<string> sourceValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                HashSet<string> matchedValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int matchedCount = 0;

                for (int index = 1; index <= context.DataRowCount; index++)
                {
                    string cellText = Convert.ToString(GetRangeValue(values, index, context.DataRowCount), CultureInfo.CurrentCulture) ?? string.Empty;
                    sourceValues.Add(cellText);
                    if (IsMatch(cellText, conditions, mode))
                    {
                        matchedCount++;
                        matchedValues.Add(cellText);
                    }
                }

                if (matchedCount == context.DataRowCount)
                {
                    ClearTargetFieldFilter(context);
                }
                else
                {
                    object[] criteria = matchedValues.Count == 0
                        ? new object[] { CreateMissingValue(sourceValues) }
                        : BuildFilterCriteria(matchedValues);
                    ApplyNativeFilter(context, criteria);
                }

                Excel.Filter filter = GetTargetFilter(context);
                currentState = new AppliedFilterState(
                    context,
                    conditions,
                    mode,
                    filter == null || !filter.On ? null : TryGetFilterSignature(filter));
                return new BatchFilterResult { CheckedCount = context.DataRowCount, MatchedCount = matchedCount };
            }
            finally
            {
                context.Application.ScreenUpdating = previousScreenUpdating;
                context.Application.EnableEvents = previousEnableEvents;
            }
        }

        internal bool Clear(BatchFilterContext context)
        {
            if (currentState == null || !currentState.Matches(context))
            {
                return false;
            }

            bool previousScreenUpdating = context.Application.ScreenUpdating;
            bool previousEnableEvents = context.Application.EnableEvents;
            try
            {
                context.Application.ScreenUpdating = false;
                context.Application.EnableEvents = false;
                ClearTargetFieldFilter(context);
                currentState = null;
                return true;
            }
            finally
            {
                context.Application.ScreenUpdating = previousScreenUpdating;
                context.Application.EnableEvents = previousEnableEvents;
            }
        }

        private static bool TryCreateListObjectContext(Excel.Application application, Excel.Worksheet worksheet, Excel.ListObject listObject, int targetColumn, out BatchFilterContext context, out string errorMessage)
        {
            context = null;
            errorMessage = null;
            if (listObject.DataBodyRange == null)
            {
                errorMessage = "当前表格没有可过滤的数据行。";
                return false;
            }

            listObject.ShowAutoFilter = true;
            int fieldIndex = targetColumn - listObject.Range.Column + 1;
            Excel.Range targetDataRange = listObject.ListColumns[fieldIndex].DataBodyRange;
            return TryBuildContext(application, worksheet, listObject.AutoFilter, listObject.Range, targetColumn, targetDataRange, out context, out errorMessage);
        }

        private static bool TryCreateFilterRangeFromSelectedHeader(Excel.Application application, Excel.Worksheet worksheet, int targetColumn, out Excel.Range filterRange, out bool cancelled, out string errorMessage)
        {
            filterRange = null;
            cancelled = false;
            errorMessage = null;
            object selection = application.InputBox("当前区域尚未启用 Excel 筛选，请选择数据的标题行。", "批量过滤", Type: 8);
            Excel.Range selectedRange = selection as Excel.Range;
            if (selectedRange == null)
            {
                cancelled = selection is bool && !(bool)selection;
                if (!cancelled)
                {
                    errorMessage = "请选择有效的标题行单元格。";
                }

                return false;
            }

            if (selectedRange.Worksheet.Name != worksheet.Name)
            {
                errorMessage = "请选择当前工作表中的标题行。";
                return false;
            }

            int firstSelectedRow = selectedRange.Row;
            int lastSelectedRow = firstSelectedRow + selectedRange.Rows.Count - 1;
            if (firstSelectedRow != lastSelectedRow)
            {
                errorMessage = "请选择单一标题行。";
                return false;
            }

            Excel.Range region = selectedRange.Cells[1, 1].CurrentRegion;
            int firstColumn = region.Column;
            int lastColumn = firstColumn + region.Columns.Count - 1;
            int lastRow = region.Row + region.Rows.Count - 1;
            if (lastRow <= firstSelectedRow)
            {
                errorMessage = "所选标题行下方没有可过滤的数据行。";
                return false;
            }

            if (targetColumn < firstColumn || targetColumn > lastColumn)
            {
                errorMessage = "最初选择的目标列不在当前数据区域内，请重新选择目标列后再试。";
                return false;
            }

            filterRange = worksheet.Range[worksheet.Cells[firstSelectedRow, firstColumn], worksheet.Cells[lastRow, lastColumn]];
            return true;
        }

        private static bool TryBuildContext(Excel.Application application, Excel.Worksheet worksheet, Excel.AutoFilter autoFilter, Excel.Range filterRange, int targetColumn, Excel.Range knownDataRange, out BatchFilterContext context, out string errorMessage)
        {
            context = null;
            errorMessage = null;
            if (!ContainsColumn(filterRange, targetColumn))
            {
                errorMessage = "最初选择的目标列不在当前数据区域内，请重新选择目标列后再试。";
                return false;
            }

            if (filterRange.Rows.Count < 2)
            {
                errorMessage = "当前筛选区域没有可过滤的数据行。";
                return false;
            }

            int fieldIndex = targetColumn - filterRange.Column + 1;
            Excel.Range dataRange = knownDataRange ?? filterRange.Columns[fieldIndex].Offset[1, 0].Resize[filterRange.Rows.Count - 1, 1];
            string headerText = Convert.ToString(filterRange.Cells[1, fieldIndex].Value2, CultureInfo.CurrentCulture);
            string columnLetter = GetColumnLetter(targetColumn);
            context = new BatchFilterContext
            {
                Application = application,
                Workbook = worksheet.Parent as Excel.Workbook,
                Worksheet = worksheet,
                AutoFilter = autoFilter,
                FilterRange = filterRange,
                DataRange = dataRange,
                ColumnDisplayName = string.IsNullOrWhiteSpace(headerText) ? columnLetter : columnLetter + " - " + headerText,
                FieldIndex = fieldIndex,
                TargetColumn = targetColumn,
                HeaderRow = filterRange.Row,
                DataRowCount = dataRange.Rows.Count
            };
            return true;
        }

        private static Excel.Filter GetTargetFilter(BatchFilterContext context)
        {
            if (context.AutoFilter == null || context.FieldIndex < 1 || context.FieldIndex > context.AutoFilter.Filters.Count)
            {
                return null;
            }

            return context.AutoFilter.Filters[context.FieldIndex];
        }

        private static string GetFilterSignature(Excel.Filter filter)
        {
            List<string> parts = new List<string>
            {
                ((int)filter.Operator).ToString(CultureInfo.InvariantCulture),
                BuildCriterionSignature(GetFilterCriterion(filter, "Criteria1"))
            };

            try
            {
                parts.Add(BuildCriterionSignature(GetFilterCriterion(filter, "Criteria2")));
            }
            catch (Exception ex) when (ex is COMException || ex is TargetInvocationException)
            {
                parts.Add(string.Empty);
            }

            return string.Join("|", parts.ToArray());
        }

        private static string TryGetFilterSignature(Excel.Filter filter)
        {
            try
            {
                return GetFilterSignature(filter);
            }
            catch (COMException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
            catch (TargetInvocationException)
            {
                return null;
            }
        }

        private static object GetFilterCriterion(Excel.Filter filter, string propertyName)
        {
            return filter.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty,
                null,
                filter,
                null,
                CultureInfo.InvariantCulture);
        }

        private static string BuildCriterionSignature(object criterion)
        {
            Array array = criterion as Array;
            if (array == null)
            {
                return BuildSignaturePart(criterion);
            }

            List<string> values = new List<string>();
            foreach (object value in array)
            {
                values.Add(BuildSignaturePart(value));
            }

            values.Sort(StringComparer.Ordinal);
            return string.Join(",", values.ToArray());
        }

        private static string BuildSignaturePart(object value)
        {
            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            return text.Length.ToString(CultureInfo.InvariantCulture) + ":" + text;
        }

        private static IList<string> ReadVisibleUniqueValues(Excel.Range dataRange)
        {
            List<string> values = new List<string>();
            HashSet<string> unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Excel.Range visibleRange;
            try
            {
                visibleRange = dataRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
            }
            catch (COMException)
            {
                return values;
            }

            foreach (Excel.Range area in visibleRange.Areas)
            {
                object areaValues = area.Value2;
                int rowCount = area.Rows.Count;
                for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    string text = Convert.ToString(GetRangeValue(areaValues, rowIndex, rowCount), CultureInfo.CurrentCulture) ?? string.Empty;
                    if (text.Length > 0 && unique.Add(text))
                    {
                        values.Add(text);
                    }
                }
            }

            return values;
        }

        private static object[] BuildFilterCriteria(IEnumerable<string> matchedValues)
        {
            List<object> criteria = new List<object>();
            foreach (string value in matchedValues)
            {
                criteria.Add(value.Length == 0 ? "=" : value);
            }

            return criteria.ToArray();
        }

        private static void ApplyNativeFilter(BatchFilterContext context, object[] criteria)
        {
            context.FilterRange.AutoFilter(context.FieldIndex, criteria, Excel.XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
        }

        private static void ClearTargetFieldFilter(BatchFilterContext context)
        {
            context.FilterRange.AutoFilter(context.FieldIndex);
        }

        private static string CreateMissingValue(ISet<string> sourceValues)
        {
            string missingValue;
            do
            {
                missingValue = "eWorkHelper_NoMatch_" + Guid.NewGuid().ToString("N");
            }
            while (sourceValues.Contains(missingValue));
            return missingValue;
        }

        private static bool IsMatch(string cellText, IList<string> conditions, BatchFilterMatchMode mode)
        {
            if (mode == BatchFilterMatchMode.NotEquals)
            {
                foreach (string condition in conditions)
                {
                    if (string.Equals(cellText, condition, StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }

                return true;
            }

            if (mode == BatchFilterMatchMode.NotContains)
            {
                foreach (string condition in conditions)
                {
                    if (cellText.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return false;
                    }
                }

                return true;
            }

            foreach (string condition in conditions)
            {
                if (mode == BatchFilterMatchMode.Equals
                    && string.Equals(cellText, condition, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (mode == BatchFilterMatchMode.Contains
                    && cellText.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static object GetRangeValue(object values, int oneBasedIndex, int rowCount)
        {
            if (rowCount == 1) return values;
            object[,] array = values as object[,];
            return array == null ? null : array[oneBasedIndex, 1];
        }

        private static bool ContainsColumn(Excel.Range range, int column)
        {
            return column >= range.Column && column < range.Column + range.Columns.Count;
        }

        private static string GetColumnLetter(int columnNumber)
        {
            string result = string.Empty;
            while (columnNumber > 0)
            {
                columnNumber--;
                result = (char)('A' + columnNumber % 26) + result;
                columnNumber /= 26;
            }
            return result;
        }

        private sealed class AppliedFilterState
        {
            private readonly Excel.Workbook workbook;
            private readonly Excel.Worksheet worksheet;
            private readonly int filterRow;
            private readonly int filterColumn;
            private readonly int filterRowCount;
            private readonly int filterColumnCount;
            private readonly int headerRow;
            private readonly int targetColumn;
            private readonly int fieldIndex;
            private readonly IList<string> conditions;
            private readonly string filterSignature;

            internal AppliedFilterState(BatchFilterContext context, IList<string> conditions, BatchFilterMatchMode matchMode, string filterSignature)
            {
                workbook = context.Workbook;
                worksheet = context.Worksheet;
                filterRow = context.FilterRange.Row;
                filterColumn = context.FilterRange.Column;
                filterRowCount = context.FilterRange.Rows.Count;
                filterColumnCount = context.FilterRange.Columns.Count;
                headerRow = context.HeaderRow;
                targetColumn = context.TargetColumn;
                fieldIndex = context.FieldIndex;
                this.conditions = new List<string>(conditions);
                MatchMode = matchMode;
                this.filterSignature = filterSignature;
            }

            internal BatchFilterMatchMode MatchMode { get; private set; }

            internal IList<string> GetConditions()
            {
                return new List<string>(conditions);
            }

            internal bool Matches(BatchFilterContext context, string currentFilterSignature)
            {
                return ReferenceEquals(workbook, context.Workbook)
                    && ReferenceEquals(worksheet, context.Worksheet)
                    && filterRow == context.FilterRange.Row
                    && filterColumn == context.FilterRange.Column
                    && filterRowCount == context.FilterRange.Rows.Count
                    && filterColumnCount == context.FilterRange.Columns.Count
                    && headerRow == context.HeaderRow
                    && targetColumn == context.TargetColumn
                    && fieldIndex == context.FieldIndex
                    && !string.IsNullOrEmpty(filterSignature)
                    && string.Equals(filterSignature, currentFilterSignature, StringComparison.Ordinal);
            }

            internal bool Matches(BatchFilterContext context)
            {
                return ReferenceEquals(workbook, context.Workbook)
                    && ReferenceEquals(worksheet, context.Worksheet)
                    && headerRow == context.HeaderRow
                    && targetColumn == context.TargetColumn
                    && fieldIndex == context.FieldIndex;
            }
        }
    }
}
