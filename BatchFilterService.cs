using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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

    /// <summary>
    /// 进度回调。返回 true 表示调用方要求取消当前操作。
    /// </summary>
    internal delegate bool BatchFilterCancellationCheck();

    /// <summary>
    /// 状态文本回调，用于向 UI 汇报长任务进度。
    /// </summary>
    internal delegate void BatchFilterProgressReport(string message);

    /// <summary>
    /// 一次批量过滤操作所需的 Excel 上下文。
    /// </summary>
    /// <remarks>
    /// 该类型拥有创建过程中产生的全部 RCW，因此实现 <see cref="IDisposable"/>；
    /// 使用方（<c>BatchFilterForm</c>）必须在窗口释放时调用 <see cref="Dispose"/>。
    /// <see cref="Application"/> 是 VSTO 宿主项（<c>Globals.ThisAddIn.Application</c>），
    /// 归运行时所有，绝不能在此释放。
    /// </remarks>
    internal sealed class BatchFilterContext : IDisposable
    {
        private readonly List<object> ownedComObjects = new List<object>();
        private bool disposed;

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

        /// <summary>登记一个在上下文生命周期内必须释放的 RCW。同一 RCW 重复登记会被忽略。</summary>
        internal void Track(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject) && !ownedComObjects.Contains(comObject))
            {
                ownedComObjects.Add(comObject);
            }
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;

            // 逆序释放，与获取顺序相反。注意绝不释放 Application（VSTO 宿主项）。
            for (int index = ownedComObjects.Count - 1; index >= 0; index--)
            {
                ComHelper.Release(ownedComObjects[index]);
            }

            ownedComObjects.Clear();
            AutoFilter = null;
            FilterRange = null;
            DataRange = null;
            Worksheet = null;
            Workbook = null;
        }
    }

    internal sealed class BatchFilterResult
    {
        /// <summary>本次检查的数据行总数。</summary>
        internal int CheckedCount { get; set; }

        /// <summary>应用筛选后工作表中实际可见的数据行数（-1 表示未统计）。</summary>
        internal int VisibleCount { get; set; } = -1;

        /// <summary>为 true 时表示 <see cref="CheckedCount"/>/<see cref="MatchedRowCount"/> 有意义，可由调用方展示计数。</summary>
        internal bool HasCounts { get; set; }

        /// <summary>
        /// 由本工具本次条件匹配到的数据行数（仅针对目标列）。
        /// 与 <see cref="VisibleCount"/> 可能不同，因为其它字段的筛选会与本次筛选取交集。
        /// </summary>
        internal int MatchedRowCount { get; set; }

        /// <summary>匹配集合超出 Excel 值列表容量或无法安全应用时为 true，此时未改动工作表。</summary>
        internal bool Applied { get; set; } = true;

        /// <summary>未应用筛选时的说明文本。</summary>
        internal string Message { get; set; }
    }

    internal sealed class BatchFilterInitialState
    {
        internal IList<string> Conditions { get; set; }
        internal BatchFilterMatchMode MatchMode { get; set; }
        internal string StatusMessage { get; set; }
    }

    internal sealed class BatchFilterService
    {
        /// <summary>单个条件长度上限（Excel 条件字符串上限为 8192 字符）。</summary>
        private const int MaxConditionLength = 8192;

        /// <summary>xlFilterValues 值列表容量上限。超过后 Excel 会直接抛出 COM 错误。</summary>
        private const int MaxFilterValueCount = 10000;

        /// <summary>
        /// 判断单元格显示文本是否可用。列宽不足时 <c>Range.Text</c> 会返回纯 "#" 串。
        /// </summary>
        private static readonly Regex OverflowTextPattern = new Regex("^#+$", RegexOptions.Compiled);

        private AppliedFilterState currentState;

        internal bool TryCreateContext(Excel.Application application, out BatchFilterContext context, out string errorMessage)
        {
            context = null;
            errorMessage = null;
            bool handedOff = false;
            try
            {
                if (application == null)
                {
                    errorMessage = "无法获取当前 Excel 应用程序。";
                    return false;
                }

                Excel.Workbooks workbooks = null;
                try
                {
                    workbooks = application.Workbooks;
                    if (workbooks == null || workbooks.Count == 0)
                    {
                        errorMessage = "没有可用的 Excel 工作簿。";
                        return false;
                    }
                }
                finally
                {
                    ComHelper.Release(workbooks);
                }

                Excel.Worksheet worksheet = null;
                Excel.Range activeCell = null;
                try
                {
                    worksheet = application.ActiveSheet as Excel.Worksheet;
                    activeCell = application.ActiveCell as Excel.Range;
                    if (worksheet == null || activeCell == null)
                    {
                        errorMessage = "没有有效的工作表或活动单元格。";
                        return false;
                    }

                    int targetColumn = activeCell.Column;
                    Excel.ListObject listObject = null;
                    try
                    {
                        listObject = activeCell.ListObject;
                        if (listObject != null)
                        {
                            return TryCreateListObjectContext(application, worksheet, listObject, targetColumn, out context, out errorMessage);
                        }
                    }
                    finally
                    {
                        ComHelper.Release(listObject);
                    }

                    bool autoFilterMode = worksheet.AutoFilterMode;
                    if (autoFilterMode)
                    {
                        Excel.AutoFilter autoFilter = null;
                        Excel.Range existingRange = null;
                        try
                        {
                            autoFilter = worksheet.AutoFilter;
                            existingRange = autoFilter == null ? null : autoFilter.Range;
                            if (existingRange == null || !ContainsColumn(existingRange, targetColumn))
                            {
                                errorMessage = "当前选择的列不在现有筛选区域内，请选择筛选区域中的列后重试。";
                                return false;
                            }

                            string builtError;
                            if (TryBuildContext(application, worksheet, autoFilter, existingRange, targetColumn, null, out context, out builtError))
                            {
                                // 上下文接管 autoFilter/existingRange 的所有权。
                                handedOff = true;
                                return true;
                            }

                            errorMessage = builtError;
                            return false;
                        }
                        finally
                        {
                            if (!handedOff)
                            {
                                ComHelper.Release(existingRange);
                                ComHelper.Release(autoFilter);
                            }
                        }
                    }

                    Excel.Range newFilterRange = null;
                    bool cancelled;
                    if (!TryCreateFilterRangeFromSelectedHeader(application, worksheet, targetColumn, out newFilterRange, out cancelled, out errorMessage))
                    {
                        return false;
                    }

                    if (cancelled)
                    {
                        return false;
                    }

                    // E-07：新建 AutoFilter 后若上下文构建失败，必须把工作表恢复原状。
                    bool guardPreviousScreenUpdating = false;
                    bool previousScreenUpdating = true;
                    try
                    {
                        try
                        {
                            previousScreenUpdating = application.ScreenUpdating;
                            guardPreviousScreenUpdating = true;
                            application.ScreenUpdating = false;
                        }
                        catch (COMException)
                        {
                            guardPreviousScreenUpdating = false;
                        }

                        try
                        {
                            newFilterRange.AutoFilter();
                            Excel.AutoFilter newAutoFilter = worksheet.AutoFilter;
                            string builtError;
                            if (!TryBuildContext(application, worksheet, newAutoFilter, newFilterRange, targetColumn, null, out context, out builtError))
                            {
                                ComHelper.Release(newAutoFilter);

                                // 建立失败：撤销刚刚新建的 AutoFilter，避免留下工作表变更。
                                RollbackSyntheticAutoFilter(newFilterRange);
                                errorMessage = builtError;
                                return false;
                            }

                            // 上下文接管 newAutoFilter/newFilterRange 的所有权。
                            handedOff = true;
                            return true;
                        }
                        finally
                        {
                            if (guardPreviousScreenUpdating)
                            {
                                try
                                {
                                    application.ScreenUpdating = previousScreenUpdating;
                                }
                                catch (Exception)
                                {
                                }
                            }
                        }
                    }
                    finally
                    {
                        // 成功交接后 newFilterRange 归上下文所有，由 Dispose 释放。
                        if (!handedOff)
                        {
                            ComHelper.Release(newFilterRange);
                        }
                    }
                }
                finally
                {
                    // 成功交接后，worksheet/activeCell 的所有权归上下文，由 Dispose 统一释放。
                    if (!handedOff)
                    {
                        ComHelper.Release(activeCell);
                        ComHelper.Release(worksheet);
                    }
                }
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
            finally
            {
                if (!handedOff && context != null)
                {
                    // 未交接给调用方（异常路径）：释放本次构建的上下文，避免泄漏。
                    context.Dispose();
                    context = null;
                }
            }
        }

        internal IList<string> NormalizeConditions(string input)
        {
            List<string> conditions = new List<string>();
            // E-13：去重语义必须与匹配语义一致（大小写不敏感），否则 "abc" 与 "ABC" 会被重复保留。
            HashSet<string> unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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

        /// <summary>
        /// E-12：校验用户输入的条件长度，避免把超长文本直接交给 Excel 触发原始 COM 错误。
        /// </summary>
        internal static string ValidateConditions(IList<string> conditions)
        {
            if (conditions == null)
            {
                return null;
            }

            foreach (string condition in conditions)
            {
                if (condition != null && condition.Length > MaxConditionLength)
                {
                    return string.Format(
                        CultureInfo.CurrentCulture,
                        "单个过滤条件最多支持 {0} 个字符，第 {1} 个条件长度为 {2}。请拆分或缩短条件后重试。",
                        MaxConditionLength,
                        IndexOfCondition(conditions, condition) + 1,
                        condition.Length);
                }
            }

            return null;
        }

        private static int IndexOfCondition(IList<string> conditions, string condition)
        {
            for (int index = 0; index < conditions.Count; index++)
            {
                if (ReferenceEquals(conditions[index], condition)
                    || string.Equals(conditions[index], condition, StringComparison.Ordinal))
                {
                    return index;
                }
            }

            return 0;
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
            AppliedFilterState state = currentState;
            if (state != null && state.Matches(context, filterSignature))
            {                return new BatchFilterInitialState
                {
                    Conditions = state.GetConditions(),
                    MatchMode = state.MatchMode,
                    StatusMessage = "已恢复上次批量过滤条件。"
                };
            }

            // E-06：与 Apply 统一使用“显示文本”可见性语义。
            IList<string> visibleValues = ReadVisibleUniqueValues(context.DataRange);
            return new BatchFilterInitialState
            {
                Conditions = visibleValues,
                MatchMode = BatchFilterMatchMode.Equals,
                StatusMessage = visibleValues.Count == 0
                    ? "当前筛选无可见数据。"
                    : string.Format(CultureInfo.CurrentCulture, "已加载当前筛选的 {0} 个可见唯一值。", visibleValues.Count)
            };
        }

        internal BatchFilterResult Apply(
            BatchFilterContext context,
            IList<string> conditions,
            BatchFilterMatchMode mode,
            BatchFilterCancellationCheck isCancelled,
            BatchFilterProgressReport reportProgress)
        {
            if (context == null)
            {
                throw new InvalidOperationException("批量过滤上下文不可用。");
            }

            string validationError = ValidateConditions(conditions);
            if (validationError != null)
            {
                throw new InvalidOperationException(validationError);
            }

            Excel.Application application = context.Application;
            bool previousScreenUpdating = true;
            bool haveScreenUpdating = false;
            bool previousEnableEvents = true;
            bool haveEnableEvents = false;
            Excel.XlCalculation previousCalculation = Excel.XlCalculation.xlCalculationAutomatic;
            bool haveCalculation = false;
            bool previousCursorSet = false;
            Excel.XlMousePointer previousCursor = Excel.XlMousePointer.xlDefault;

            try
            {
                try
                {
                    previousScreenUpdating = application.ScreenUpdating;
                    haveScreenUpdating = true;
                    application.ScreenUpdating = false;
                }
                catch (COMException)
                {
                }

                try
                {
                    previousEnableEvents = application.EnableEvents;
                    haveEnableEvents = true;
                    application.EnableEvents = false;
                }
                catch (COMException)
                {
                }

                // E-05：长任务期间冻结重算，避免每个筛选步骤都触发整表重算。
                try
                {
                    previousCalculation = application.Calculation;
                    haveCalculation = true;
                    application.Calculation = Excel.XlCalculation.xlCalculationManual;
                }
                catch (COMException)
                {
                }

                // E-05：等待光标与状态栏进度，均在 finally 中恢复。
                try
                {
                    previousCursor = application.Cursor;
                    previousCursorSet = true;
                    application.Cursor = Excel.XlMousePointer.xlWait;
                }
                catch (COMException)
                {
                    previousCursorSet = false;
                }

                return ApplyCore(context, conditions, mode, isCancelled, reportProgress);
            }
            finally
            {
                // E-08：每个恢复动作单独 try/catch，避免恢复失败替换掉在途异常。
                if (previousCursorSet)
                {
                    try
                    {
                        application.Cursor = previousCursor;
                    }
                    catch (Exception)
                    {
                    }
                }

                try
                {
                    application.StatusBar = false;
                }
                catch (Exception)
                {
                }

                if (haveCalculation)
                {
                    try
                    {
                        application.Calculation = previousCalculation;
                    }
                    catch (Exception)
                    {
                    }
                }

                if (haveEnableEvents)
                {
                    try
                    {
                        application.EnableEvents = previousEnableEvents;
                    }
                    catch (Exception)
                    {
                    }
                }

                if (haveScreenUpdating)
                {
                    try
                    {
                        application.ScreenUpdating = previousScreenUpdating;
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private BatchFilterResult ApplyCore(
            BatchFilterContext context,
            IList<string> conditions,
            BatchFilterMatchMode mode,
            BatchFilterCancellationCheck isCancelled,
            BatchFilterProgressReport reportProgress)
        {
            IList<CellDisplay> cells = ReadCellDisplays(context, isCancelled, reportProgress);

            List<string> matchedValues = new List<string>();
            HashSet<string> matchedUnique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int matchedCount = 0;

            foreach (CellDisplay cell in cells)
            {
                if (IsMatch(cell, conditions, mode))
                {
                    matchedCount++;
                    if (matchedUnique.Add(cell.Text))
                    {
                        matchedValues.Add(cell.Text);
                    }
                }
            }

            BatchFilterResult result = new BatchFilterResult
            {
                CheckedCount = cells.Count,
                MatchedRowCount = matchedCount,
                HasCounts = true
            };

            if (matchedCount == cells.Count)
            {
                // 全部命中：清除该字段的筛选，保持所有行可见。
                ClearTargetFieldFilter(context);
            }
            else if (matchedCount == 0)
            {
                // 无命中：不应用筛选（隐藏全部行会让用户无法回到数据），明确告知而不是留下 0 行可见。
                result.Applied = false;
                result.Message = "没有单元格与当前条件匹配，未修改工作表筛选状态。";
            }
            else
            {
                if (matchedValues.Count > MaxFilterValueCount)
                {
                    // E-12：不要把一个必然失败的超大值列表交给 Excel。
                    result.Applied = false;
                    result.Message = string.Format(
                        CultureInfo.CurrentCulture,
                        "匹配到 {0} 个唯一值，超过 Excel 筛选值列表上限（{1} 项），未修改工作表。请增加条件以缩小范围。",
                        matchedValues.Count,
                        MaxFilterValueCount);
                }
                else
                {
                    object[] criteria = BuildFilterCriteria(matchedValues);
                    ApplyNativeFilter(context, criteria);
                }
            }

            if (result.Applied)
            {
                // E-06：报告实际可见行数，而不是内存中的条件匹配数
                //（其它字段的筛选会与本次筛选取交集，两者可能不同）。
                result.VisibleCount = CountVisibleDataRows(context);
            }

            Excel.Filter filter = GetTargetFilter(context);
            bool filterOn = filter != null && filter.On;
            currentState = new AppliedFilterState(
                context,
                conditions,
                mode,
                filterOn ? TryGetFilterSignature(filter) : null);
            return result;
        }

        /// <summary>
        /// 统计目标数据列中应用筛选后实际可见的行数（返回总数后再减去被其它字段筛掉的行）。
        /// </summary>
        private static int CountVisibleDataRows(BatchFilterContext context)
        {
            Excel.Range dataRange = context.DataRange;
            if (dataRange == null)
            {
                return -1;
            }

            int visibleRows;
            Excel.Range visibleRange = null;
            try
            {
                try
                {
                    visibleRange = dataRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                }
                catch (COMException)
                {
                    // 没有可见单元格时 SpecialCells 会抛 COM 异常。
                    return 0;
                }

                if (visibleRange == null)
                {
                    return 0;
                }

                int areaCount = CountAreas(visibleRange);
                visibleRows = 0;
                for (int areaIndex = 1; areaIndex <= areaCount; areaIndex++)
                {
                    Excel.Range area = null;
                    Excel.Range areaRows = null;
                    try
                    {
                        area = GetArea(visibleRange, areaIndex);
                        if (area == null)
                        {
                            continue;
                        }

                        areaRows = area.Rows;
                        visibleRows += areaRows.Count;
                    }
                    finally
                    {
                        ComHelper.Release(areaRows);
                        ComHelper.Release(area);
                    }
                }
            }
            finally
            {
                ComHelper.Release(visibleRange);
            }

            return visibleRows;
        }

        internal bool Clear(BatchFilterContext context)
        {
            if (context == null)
            {
                return false;
            }

            AppliedFilterState state = currentState;
            if (state == null || !state.Matches(context))
            {
                return false;
            }

            Excel.Application application = context.Application;
            bool previousScreenUpdating = true;
            bool haveScreenUpdating = false;
            bool previousEnableEvents = true;
            bool haveEnableEvents = false;
            bool previousCursorSet = false;
            Excel.XlMousePointer previousCursor = Excel.XlMousePointer.xlDefault;

            try
            {
                try
                {
                    previousScreenUpdating = application.ScreenUpdating;
                    haveScreenUpdating = true;
                    application.ScreenUpdating = false;
                }
                catch (COMException)
                {
                }

                try
                {
                    previousEnableEvents = application.EnableEvents;
                    haveEnableEvents = true;
                    application.EnableEvents = false;
                }
                catch (COMException)
                {
                }

                try
                {
                    previousCursor = application.Cursor;
                    previousCursorSet = true;
                    application.Cursor = Excel.XlMousePointer.xlWait;
                }
                catch (COMException)
                {
                    previousCursorSet = false;
                }

                ClearTargetFieldFilter(context);
                currentState = null;
                return true;
            }
            finally
            {
                if (previousCursorSet)
                {
                    try
                    {
                        application.Cursor = previousCursor;
                    }
                    catch (Exception)
                    {
                    }
                }

                try
                {
                    application.StatusBar = false;
                }
                catch (Exception)
                {
                }

                if (haveEnableEvents)
                {
                    try
                    {
                        application.EnableEvents = previousEnableEvents;
                    }
                    catch (Exception)
                    {
                    }
                }

                if (haveScreenUpdating)
                {
                    try
                    {
                        application.ScreenUpdating = previousScreenUpdating;
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static bool TryCreateListObjectContext(
            Excel.Application application,
            Excel.Worksheet worksheet,
            Excel.ListObject listObject,
            int targetColumn,
            out BatchFilterContext context,
            out string errorMessage)
        {
            context = null;
            errorMessage = null;

            Excel.Range listObjectRange = null;
            Excel.Range dataBodyRange = null;
            bool previousShowAutoFilter = false;
            bool haveShowAutoFilter = false;
            bool changedShowAutoFilter = false;
            try
            {
                // E-07：先记录 ShowAutoFilter 原值，失败时回滚。
                try
                {
                    previousShowAutoFilter = listObject.ShowAutoFilter;
                    haveShowAutoFilter = true;
                }
                catch (COMException)
                {
                    haveShowAutoFilter = false;
                }

                listObjectRange = listObject.Range;
                dataBodyRange = listObject.DataBodyRange;
                if (dataBodyRange == null)
                {
                    errorMessage = "当前表格没有可过滤的数据行。";
                    return false;
                }

                if (haveShowAutoFilter && !previousShowAutoFilter)
                {
                    listObject.ShowAutoFilter = true;
                    changedShowAutoFilter = true;
                }

                int fieldIndex = targetColumn - listObjectRange.Column + 1;
                Excel.ListColumn listColumn = null;
                Excel.Range targetDataRange = null;
                Excel.AutoFilter listAutoFilter = null;
                try
                {
                    listColumn = listObject.ListColumns[fieldIndex];
                    targetDataRange = listColumn.DataBodyRange;
                    listAutoFilter = listObject.AutoFilter;

                    bool built = TryBuildContext(
                        application,
                        worksheet,
                        listAutoFilter,
                        listObjectRange,
                        targetColumn,
                        targetDataRange,
                        out context,
                        out errorMessage);

                    if (!built && changedShowAutoFilter)
                    {
                        // 建立失败：恢复原来的 ShowAutoFilter 状态。
                        try
                        {
                            listObject.ShowAutoFilter = previousShowAutoFilter;
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return built;
                }
                finally
                {
                    // 这些中间 RCW 一旦交给上下文就由上下文负责释放；未交接的在这里释放。
                    if (context == null)
                    {
                        ComHelper.Release(listAutoFilter);
                        ComHelper.Release(targetDataRange);
                    }

                    ComHelper.Release(listColumn);
                }
            }
            catch (COMException)
            {
                errorMessage = "无法识别表格筛选区域，请确认工作表未受保护并重试。";
                return false;
            }
            finally
            {
                // 成功交接后，listObjectRange/dataBodyRange 的所有权归上下文，由 Dispose 统一释放。
                if (context == null)
                {
                    ComHelper.Release(dataBodyRange);
                    ComHelper.Release(listObjectRange);
                }
            }
        }

        private static bool TryCreateFilterRangeFromSelectedHeader(
            Excel.Application application,
            Excel.Worksheet worksheet,
            int targetColumn,
            out Excel.Range filterRange,
            out bool cancelled,
            out string errorMessage)
        {
            filterRange = null;
            cancelled = false;
            errorMessage = null;

            object selection = null;
            Excel.Range selectedRange = null;
            Excel.Range firstCell = null;
            Excel.Range region = null;
            Excel.Range headerCell = null;
            try
            {
                selection = application.InputBox("当前区域尚未启用 Excel 筛选，请选择数据的标题行。", "批量过滤", Type: 8);
                selectedRange = selection as Excel.Range;
                if (selectedRange == null)
                {
                    cancelled = selection is bool && !(bool)selection;
                    if (!cancelled)
                    {
                        errorMessage = "请选择有效的标题行单元格。";
                    }

                    return false;
                }

                // E-14：工作表校验必须同时比较工作簿与工作表，避免同名工作表误判。
                Excel.Worksheet selectedWorksheet = null;
                try
                {
                    selectedWorksheet = selectedRange.Worksheet as Excel.Worksheet;
                }
                catch (COMException)
                {
                    selectedWorksheet = null;
                }

                bool sameWorksheet;
                try
                {
                    sameWorksheet = IsSameWorksheet(selectedWorksheet, worksheet);
                }
                finally
                {
                    ComHelper.Release(selectedWorksheet);
                }

                if (!sameWorksheet)
                {
                    errorMessage = "请选择当前工作表中的标题行。";
                    return false;
                }

                int firstSelectedRow = selectedRange.Row;
                int selectedRowCount;
                Excel.Range selectedRows = null;
                try
                {
                    selectedRows = selectedRange.Rows;
                    selectedRowCount = selectedRows.Count;
                }
                finally
                {
                    ComHelper.Release(selectedRows);
                }

                int lastSelectedRow = firstSelectedRow + selectedRowCount - 1;
                if (firstSelectedRow != lastSelectedRow)
                {
                    errorMessage = "请选择单一标题行。";
                    return false;
                }

                firstCell = selectedRange.Cells[1, 1] as Excel.Range;
                if (firstCell == null)
                {
                    errorMessage = "请选择有效的标题行单元格。";
                    return false;
                }

                region = firstCell.CurrentRegion;
                if (region == null)
                {
                    errorMessage = "无法识别所选标题行所在的数据区域。";
                    return false;
                }

                int firstColumn = region.Column;
                int lastColumn;
                Excel.Range regionColumns = null;
                try
                {
                    regionColumns = region.Columns;
                    lastColumn = firstColumn + regionColumns.Count - 1;
                }
                finally
                {
                    ComHelper.Release(regionColumns);
                }

                int lastRow;
                Excel.Range regionRows = null;
                try
                {
                    regionRows = region.Rows;
                    lastRow = region.Row + regionRows.Count - 1;
                }
                finally
                {
                    ComHelper.Release(regionRows);
                }

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

                Excel.Range firstCorner = null;
                Excel.Range lastCorner = null;
                try
                {
                    firstCorner = worksheet.Cells[firstSelectedRow, firstColumn];
                    lastCorner = worksheet.Cells[lastRow, lastColumn];
                    filterRange = worksheet.Range[firstCorner, lastCorner];
                }
                finally
                {
                    ComHelper.Release(lastCorner);
                    ComHelper.Release(firstCorner);
                }

                return filterRange != null;
            }
            catch (COMException)
            {
                errorMessage = "无法识别所选标题行的数据区域，请重试。";
                return false;
            }
            finally
            {
                ComHelper.Release(headerCell);
                ComHelper.Release(region);
                ComHelper.Release(firstCell);
                ComHelper.Release(selectedRange);
                ComHelper.Release(selection);
            }
        }

        private static bool TryBuildContext(
            Excel.Application application,
            Excel.Worksheet worksheet,
            Excel.AutoFilter autoFilter,
            Excel.Range filterRange,
            int targetColumn,
            Excel.Range knownDataRange,
            out BatchFilterContext context,
            out string errorMessage)
        {
            context = null;
            errorMessage = null;

            if (!ContainsColumn(filterRange, targetColumn))
            {
                errorMessage = "最初选择的目标列不在当前数据区域内，请重新选择目标列后再试。";
                return false;
            }

            int rowCount;
            Excel.Range filterRows = null;
            try
            {
                filterRows = filterRange.Rows;
                rowCount = filterRows.Count;
            }
            finally
            {
                ComHelper.Release(filterRows);
            }

            if (rowCount < 2)
            {
                errorMessage = "当前筛选区域没有可过滤的数据行。";
                return false;
            }

            int firstColumn = filterRange.Column;
            int firstRow = filterRange.Row;
            int fieldIndex = targetColumn - firstColumn + 1;

            Excel.Range columns = null;
            Excel.Range columnSlice = null;
            Excel.Range offsetRange = null;
            Excel.Range dataRange = null;
            Excel.Range headerRange = null;
            Excel.Range headerCell = null;
            Excel.Workbook workbook = null;
            try
            {
                if (knownDataRange == null)
                {
                    columns = filterRange.Columns;
                    columnSlice = columns[fieldIndex];
                    offsetRange = columnSlice.Offset[1, 0];
                    dataRange = offsetRange.Resize[rowCount - 1, 1];
                }
                else
                {
                    dataRange = knownDataRange;
                }

                headerRange = filterRange.Cells[1, fieldIndex] as Excel.Range;
                headerCell = headerRange;
                string headerText = headerCell == null
                    ? null
                    : Convert.ToString(headerCell.Value2, CultureInfo.CurrentCulture);

                int dataRowCount;
                Excel.Range dataRows = null;
                try
                {
                    dataRows = dataRange.Rows;
                    dataRowCount = dataRows.Count;
                }
                finally
                {
                    ComHelper.Release(dataRows);
                }

                workbook = worksheet.Parent as Excel.Workbook;

                BatchFilterContext built = new BatchFilterContext
                {
                    Application = application,
                    Workbook = workbook,
                    Worksheet = worksheet,
                    AutoFilter = autoFilter,
                    FilterRange = filterRange,
                    DataRange = dataRange,
                    ColumnDisplayName = string.IsNullOrWhiteSpace(headerText)
                        ? GetColumnLetter(targetColumn)
                        : GetColumnLetter(targetColumn) + " - " + headerText,
                    FieldIndex = fieldIndex,
                    TargetColumn = targetColumn,
                    HeaderRow = firstRow,
                    DataRowCount = dataRowCount
                };

                // 上下文负责在释放时回收这些 RCW。
                built.Track(headerRange);
                built.Track(workbook);
                built.Track(dataRange);
                built.Track(autoFilter);
                built.Track(filterRange);
                built.Track(worksheet);

                context = built;
                return true;
            }
            catch (COMException)
            {
                errorMessage = "无法识别筛选区域结构，请确认工作表未受保护并重试。";
                return false;
            }
            finally
            {
                if (context == null)
                {
                    // 未交接给调用方的中间对象在这里释放。
                    ComHelper.Release(headerCell);
                    ComHelper.Release(headerRange);
                    ComHelper.Release(workbook);
                    if (knownDataRange == null)
                    {
                        ComHelper.Release(dataRange);
                    }

                    ComHelper.Release(offsetRange);
                    ComHelper.Release(columnSlice);
                    ComHelper.Release(columns);
                }
            }
        }

        private static Excel.Filter GetTargetFilter(BatchFilterContext context)
        {
            if (context.AutoFilter == null || context.FieldIndex < 1)
            {
                return null;
            }

            try
            {
                // E-04：Filters 访问在 AutoFilter 被外部改动/工作表受保护时可能抛 COM 异常，
                // 这里守护住，避免异常越过 Ribbon 边界。
                Excel.Filters filters = context.AutoFilter.Filters;
                if (filters == null)
                {
                    return null;
                }

                if (context.FieldIndex > filters.Count)
                {
                    return null;
                }

                Excel.Filter filter = filters[context.FieldIndex];
                ComHelper.Release(filters);
                return filter;
            }
            catch (COMException)
            {
                return null;
            }
            catch (InvalidCastException)
            {
                return null;
            }
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

        /// <summary>
        /// E-02 / E-06：读取只显示文本（<c>Range.Text</c>），与用户看到的单元格内容和 Excel
        /// AutoFilter 的值列表语义保持一致；日期不再退化为 OADate 序列号，百分比不再退化为原始小数。
        /// </summary>
        private static IList<string> ReadVisibleUniqueValues(Excel.Range dataRange)
        {
            List<string> values = new List<string>();
            HashSet<string> unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (dataRange == null)
            {
                return values;
            }

            Excel.Range visibleRange = null;
            try
            {
                try
                {
                    visibleRange = dataRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                }
                catch (COMException)
                {
                    return values;
                }

                if (visibleRange == null)
                {
                    return values;
                }

                int areaCount = CountAreas(visibleRange);
                for (int areaIndex = 1; areaIndex <= areaCount; areaIndex++)
                {
                    Excel.Range area = null;
                    try
                    {
                        area = GetArea(visibleRange, areaIndex);
                        if (area == null)
                        {
                            continue;
                        }

                        IList<string> areaTexts = ReadDisplayTexts(area);
                        foreach (string text in areaTexts)
                        {
                            if (text.Length > 0 && unique.Add(text))
                            {
                                values.Add(text);
                            }
                        }
                    }
                    finally
                    {
                        ComHelper.Release(area);
                    }
                }
            }
            finally
            {
                ComHelper.Release(visibleRange);
            }

            return values;
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

        /// <summary>
        /// 读取一个单列区域的全部单元格显示文本。返回顺序与区域行顺序一致。
        /// </summary>
        private static IList<string> ReadDisplayTexts(Excel.Range range)
        {
            List<string> texts = new List<string>();
            int rowCount;
            Excel.Range rows = null;
            try
            {
                rows = range.Rows;
                rowCount = rows.Count;
            }
            finally
            {
                ComHelper.Release(rows);
            }

            object textValues = null;
            object rawValues = null;
            try
            {
                try
                {
                    textValues = range.Text;
                }
                catch (COMException)
                {
                    textValues = null;
                }

                try
                {
                    rawValues = range.Value2;
                }
                catch (COMException)
                {
                    rawValues = null;
                }

                for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    texts.Add(ResolveDisplayText(
                        GetMatrixValue(textValues, rowIndex, rowCount),
                        GetMatrixValue(rawValues, rowIndex, rowCount)));
                }
            }
            finally
            {
                ComHelper.Release(rawValues);
                ComHelper.Release(textValues);
            }

            return texts;
        }

        /// <summary>
        /// 读取目标列全部单元格的显示文本与辅助数值，并支持协作取消与进度汇报。
        /// </summary>
        private static IList<CellDisplay> ReadCellDisplays(
            BatchFilterContext context,
            BatchFilterCancellationCheck isCancelled,
            BatchFilterProgressReport reportProgress)
        {
            List<CellDisplay> displays = new List<CellDisplay>(context.DataRowCount);
            Excel.Range dataRange = context.DataRange;
            if (dataRange == null)
            {
                return displays;
            }

            Excel.Range rows = null;
            object textValues = null;
            object rawValues = null;
            try
            {
                int rowCount;
                rows = dataRange.Rows;
                rowCount = rows.Count;

                try
                {
                    textValues = dataRange.Text;
                }
                catch (COMException)
                {
                    textValues = null;
                }

                try
                {
                    rawValues = dataRange.Value2;
                }
                catch (COMException)
                {
                    rawValues = null;
                }

                int reportInterval = Math.Max(1, rowCount / 100);
                for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
                {
                    string text = ResolveDisplayText(
                        GetMatrixValue(textValues, rowIndex, rowCount),
                        GetMatrixValue(rawValues, rowIndex, rowCount));

                    displays.Add(new CellDisplay
                    {
                        Text = text,
                        AlternateText = BuildAlternateText(text, GetMatrixValue(rawValues, rowIndex, rowCount))
                    });

                    if (isCancelled != null && isCancelled())
                    {
                        throw new OperationCanceledException("批量过滤已取消。");
                    }

                    if (reportProgress != null && (rowIndex % reportInterval == 0))
                    {
                        reportProgress(string.Format(
                            CultureInfo.CurrentCulture,
                            "正在读取数据 {0}/{1} … 按 Esc 取消",
                            rowIndex,
                            rowCount));
                    }
                }
            }
            finally
            {
                ComHelper.Release(rawValues);
                ComHelper.Release(textValues);
                ComHelper.Release(rows);
            }

            return displays;
        }

        /// <summary>
        /// 解析单元格显示文本：优先使用 <c>Range.Text</c>；列宽不足返回 "#" 串时回退到格式化后的值，
        /// 避免把纯 "#" 当作真实内容参与匹配与值列表。
        /// </summary>
        private static string ResolveDisplayText(object textValue, object rawValue)
        {
            string text = textValue as string ?? (textValue == null ? null : Convert.ToString(textValue, CultureInfo.CurrentCulture));
            if (!string.IsNullOrEmpty(text) && !OverflowTextPattern.IsMatch(text.Trim()))
            {
                return text;
            }

            return FallbackText(rawValue);
        }

        private static string FallbackText(object rawValue)
        {
            if (rawValue == null)
            {
                return string.Empty;
            }

            if (rawValue is string rawString)
            {
                return rawString;
            }

            if (rawValue is DateTime dateValue)
            {
                return dateValue.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }

            if (rawValue is bool boolValue)
            {
                return boolValue ? "TRUE" : "FALSE";
            }

            if (rawValue is double || rawValue is float || rawValue is decimal
                || rawValue is int || rawValue is long || rawValue is short || rawValue is byte)
            {
                // General 格式下 Excel 使用 invariant 数字文本。
                return Convert.ToString(rawValue, CultureInfo.InvariantCulture) ?? string.Empty;
            }

            try
            {
                return Convert.ToString(rawValue, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch (InvalidCastException)
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 构造辅助匹配候选：百分比显示 "15%" 时提供 "0.15"；显示文本与原始数值文本不同时提供后者，
        /// 以便用户以原始数值或不受区域格式影响的写法命中同一个单元格。
        /// </summary>
        private static string BuildAlternateText(string displayText, object rawValue)
        {
            if (string.IsNullOrEmpty(displayText))
            {
                return null;
            }

            if (displayText.EndsWith("%", StringComparison.Ordinal))
            {
                string numberPart = displayText.Substring(0, displayText.Length - 1).Trim();
                double percentValue;
                if (double.TryParse(numberPart, NumberStyles.Any, CultureInfo.CurrentCulture, out percentValue)
                    || double.TryParse(numberPart, NumberStyles.Any, CultureInfo.InvariantCulture, out percentValue))
                {
                    return (percentValue / 100d).ToString(CultureInfo.InvariantCulture);
                }

                return null;
            }

            if (!(rawValue is double || rawValue is float || rawValue is decimal
                || rawValue is int || rawValue is long || rawValue is short || rawValue is byte))
            {
                return null;
            }

            string rawText = Convert.ToString(rawValue, CultureInfo.InvariantCulture);
            return string.Equals(rawText, displayText, StringComparison.Ordinal) ? null : rawText;
        }

        /// <summary>
        /// 从 <c>Range.Text</c> / <c>Range.Value2</c> 的返回值中取出第 N 行的值。
        /// 单行区域会返回标量而不是二维数组。
        /// </summary>
        private static object GetMatrixValue(object values, int oneBasedIndex, int rowCount)
        {
            if (values == null)
            {
                return null;
            }

            if (rowCount == 1)
            {
                return values;
            }

            object[,] array = values as object[,];
            if (array != null)
            {
                return array[oneBasedIndex, 1];
            }

            string[] flat = values as string[];
            if (flat != null && oneBasedIndex - 1 < flat.Length)
            {
                return flat[oneBasedIndex - 1];
            }

            return null;
        }

        /// <summary>
        /// E-12：转义 AutoFilter 通配符字符，使字面量条件不会被当作通配模式。
        /// </summary>
        private static string EscapeWildcardCharacters(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "=";
            }

            StringBuilder builder = new StringBuilder(value.Length);
            foreach (char character in value)
            {
                if (character == '~' || character == '*' || character == '?')
                {
                    builder.Append('~');
                }

                builder.Append(character);
            }

            return builder.ToString();
        }

        private static object[] BuildFilterCriteria(IEnumerable<string> matchedValues)
        {
            List<object> criteria = new List<object>();
            foreach (string value in matchedValues)
            {
                criteria.Add(EscapeWildcardCharacters(value));
            }

            return criteria.ToArray();
        }

        private static void ApplyNativeFilter(BatchFilterContext context, object[] criteria)
        {
            context.FilterRange.AutoFilter(
                context.FieldIndex,
                criteria,
                Excel.XlAutoFilterOperator.xlFilterValues,
                Type.Missing,
                true);
        }

        private static void ClearTargetFieldFilter(BatchFilterContext context)
        {
            context.FilterRange.AutoFilter(context.FieldIndex);
        }

        /// <summary>E-07：建好工作表后又失败时，撤掉刚刚新建的 AutoFilter。</summary>
        private static void RollbackSyntheticAutoFilter(Excel.Range filterRange)
        {
            if (filterRange == null)
            {
                return;
            }

            try
            {
                // 无参调用会移除该区域上的 AutoFilter（等价于关闭 AutoFilterMode）。
                filterRange.AutoFilter();
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// 判断单元格是否命中条件。主匹配对象是显示文本（与文档承诺一致）；
        /// 百分比/区域格式差异时额外比较辅助候选（如 "15%" 与 "0.15"）。
        /// </summary>
        private static bool IsMatch(CellDisplay cell, IList<string> conditions, BatchFilterMatchMode mode)
        {
            bool matched = IsMatch(cell.Text, conditions, mode);
            if (!matched && !string.IsNullOrEmpty(cell.AlternateText))
            {
                matched = IsMatch(cell.AlternateText, conditions, mode);
            }

            return matched;
        }

        private static bool IsMatch(string cellText, IList<string> conditions, BatchFilterMatchMode mode)
        {
            string text = cellText ?? string.Empty;
            if (mode == BatchFilterMatchMode.NotEquals)
            {
                foreach (string condition in conditions)
                {
                    if (string.Equals(text, condition, StringComparison.OrdinalIgnoreCase))
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
                    if (text.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return false;
                    }
                }

                return true;
            }

            foreach (string condition in conditions)
            {
                if (mode == BatchFilterMatchMode.Equals
                    && string.Equals(text, condition, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (mode == BatchFilterMatchMode.Contains
                    && text.IndexOf(condition, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool ContainsColumn(Excel.Range range, int column)
        {
            int firstColumn = range.Column;
            int columnCount;
            Excel.Range columns = null;
            try
            {
                columns = range.Columns;
                columnCount = columns.Count;
            }
            finally
            {
                ComHelper.Release(columns);
            }

            return column >= firstColumn && column < firstColumn + columnCount;
        }

        /// <summary>E-14：同时校验工作簿与工作表，避免同名工作表被误认为目标表。</summary>
        private static bool IsSameWorksheet(Excel.Worksheet candidate, Excel.Worksheet expected)
        {
            if (candidate == null || expected == null)
            {
                return false;
            }

            string candidateName = candidate.Name;
            string expectedName = expected.Name;
            if (!string.Equals(candidateName, expectedName, StringComparison.Ordinal))
            {
                return false;
            }

            Excel.Workbook candidateWorkbook = null;
            Excel.Workbook expectedWorkbook = null;
            try
            {
                candidateWorkbook = candidate.Parent as Excel.Workbook;
                expectedWorkbook = expected.Parent as Excel.Workbook;
                if (candidateWorkbook == null || expectedWorkbook == null)
                {
                    return false;
                }

                return string.Equals(
                    candidateWorkbook.FullName,
                    expectedWorkbook.FullName,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                ComHelper.Release(expectedWorkbook);
                ComHelper.Release(candidateWorkbook);
            }
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

        /// <summary>
        /// 单元格的显示文本与辅助匹配候选。
        /// </summary>
        private struct CellDisplay
        {
            /// <summary>单元格显示文本（与 Excel 值列表语义一致）。</summary>
            internal string Text;

            /// <summary>
            /// 辅助匹配候选：百分比单元格显示 "15%" 时提供 "0.15"，
            /// 千分位等区域格式差异时提供原始数值文本。
            /// </summary>
            internal string AlternateText;
        }

        /// <summary>
        /// E-03：只保存身份信息（工作簿全名、工作表名、行列号），不再缓存 Workbook/Worksheet RCW。
        /// </summary>
        private sealed class AppliedFilterState
        {
            private readonly string workbookFullName;
            private readonly string worksheetName;
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
                Excel.Workbook workbook = context.Workbook;
                Excel.Worksheet worksheet = context.Worksheet;
                Excel.Range filterRange = context.FilterRange;

                workbookFullName = workbook == null ? null : workbook.FullName;
                worksheetName = worksheet == null ? null : worksheet.Name;
                filterRow = filterRange == null ? 0 : filterRange.Row;
                filterColumn = filterRange == null ? 0 : filterRange.Column;
                filterRowCount = context.FilterRange == null ? 0 : CountOf(filterRange, true);
                filterColumnCount = context.FilterRange == null ? 0 : CountOf(filterRange, false);
                headerRow = context.HeaderRow;
                targetColumn = context.TargetColumn;
                fieldIndex = context.FieldIndex;
                this.conditions = new List<string>(conditions);
                MatchMode = matchMode;
                this.filterSignature = filterSignature;
            }

            internal BatchFilterMatchMode MatchMode { get; private set; }

            private static int CountOf(Excel.Range range, bool rows)
            {
                Excel.Range sub = null;
                try
                {
                    sub = rows ? range.Rows : range.Columns;
                    return sub == null ? 0 : sub.Count;
                }
                catch (COMException)
                {
                    return 0;
                }
                finally
                {
                    ComHelper.Release(sub);
                }
            }

            internal IList<string> GetConditions()
            {
                return new List<string>(conditions);
            }

            internal bool Matches(BatchFilterContext context, string currentFilterSignature)
            {
                if (context == null)
                {
                    return false;
                }

                Excel.Workbook workbook = context.Workbook;
                Excel.Worksheet worksheet = context.Worksheet;
                Excel.Range filterRange = context.FilterRange;
                if (workbook == null || worksheet == null || filterRange == null)
                {
                    return false;
                }

                try
                {
                    return string.Equals(workbookFullName, workbook.FullName, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(worksheetName, worksheet.Name, StringComparison.Ordinal)
                        && filterRow == filterRange.Row
                        && filterColumn == filterRange.Column
                        && filterRowCount == CountOf(filterRange, true)
                        && filterColumnCount == CountOf(filterRange, false)
                        && headerRow == context.HeaderRow
                        && targetColumn == context.TargetColumn
                        && fieldIndex == context.FieldIndex
                        && !string.IsNullOrEmpty(filterSignature)
                        && string.Equals(filterSignature, currentFilterSignature, StringComparison.Ordinal);
                }
                catch (COMException)
                {
                    // 之前记录的工作簿/工作表可能已被关闭（RPC_E_DISCONNECTED）：按不匹配处理。
                    return false;
                }
                catch (InvalidCastException)
                {
                    return false;
                }
            }

            internal bool Matches(BatchFilterContext context)
            {
                if (context == null)
                {
                    return false;
                }

                Excel.Workbook workbook = context.Workbook;
                Excel.Worksheet worksheet = context.Worksheet;
                if (workbook == null || worksheet == null)
                {
                    return false;
                }

                try
                {
                    return string.Equals(workbookFullName, workbook.FullName, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(worksheetName, worksheet.Name, StringComparison.Ordinal)
                        && headerRow == context.HeaderRow
                        && targetColumn == context.TargetColumn
                        && fieldIndex == context.FieldIndex;
                }
                catch (COMException)
                {
                    return false;
                }
                catch (InvalidCastException)
                {
                    return false;
                }
            }
        }
    }
}
