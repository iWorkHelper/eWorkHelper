using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace eWorkhelper
{
    internal enum ExcelDiffOperationStatus
    {
        Completed,
        Cancelled,
        Failed
    }

    internal sealed class ExcelDiffLog
    {
        private readonly IList<string> lines = new List<string>();

        internal void Add(string message)
        {
            if (!string.IsNullOrEmpty(message))
            {
                lines.Add(message);
            }
        }

        internal IList<string> Lines { get { return lines; } }

        public override string ToString()
        {
            return string.Join(Environment.NewLine, lines);
        }
    }

    internal sealed class ExcelDiffWorkbookSource
    {
        private ExcelDiffWorkbookSource(string filePath, Excel.Workbook workbook)
        {
            FilePath = filePath;
            Workbook = workbook;
        }

        internal string FilePath { get; private set; }
        internal Excel.Workbook Workbook { get; private set; }

        internal static ExcelDiffWorkbookSource FromFile(string filePath)
        {
            return new ExcelDiffWorkbookSource(filePath, null);
        }

        internal static ExcelDiffWorkbookSource FromOpenWorkbook(Excel.Workbook workbook)
        {
            return new ExcelDiffWorkbookSource(null, workbook);
        }
    }

    internal sealed class ExcelDiffSheetPair
    {
        internal ExcelDiffSheetPair(string baselineSheetName, string candidateSheetName, string displayName)
        {
            BaselineSheetName = baselineSheetName ?? string.Empty;
            CandidateSheetName = candidateSheetName ?? string.Empty;
            DisplayName = string.IsNullOrEmpty(displayName)
                ? string.Format("【新】{0} <<【旧】{1}", CandidateSheetName, BaselineSheetName)
                : displayName;
        }

        internal string BaselineSheetName { get; private set; }
        internal string CandidateSheetName { get; private set; }
        internal string DisplayName { get; private set; }
    }

    internal sealed class ExcelDiffOptions
    {
        internal ExcelDiffOptions()
        {
            HighlightDifferences = true;
            AddBaselineComment = false;
            HighlightColor = Color.Red;
            CommentAuthor = "Diff Excel";
            OverwriteExistingOutput = false;
            ValidateOutputAfterSave = true;
        }

        internal bool HighlightDifferences { get; set; }
        internal bool AddBaselineComment { get; set; }
        internal Color HighlightColor { get; set; }
        internal string CommentAuthor { get; set; }
        internal bool OverwriteExistingOutput { get; set; }
        internal bool ValidateOutputAfterSave { get; set; }
    }

    internal sealed class ExcelDiffOutputOptions
    {
        internal ExcelDiffOutputOptions(string outputPath)
        {
            OutputPath = outputPath;
        }

        internal string OutputPath { get; private set; }
    }

    internal sealed class ExcelDiffRequest
    {
        internal ExcelDiffRequest(
            ExcelDiffWorkbookSource baseline,
            ExcelDiffWorkbookSource candidate,
            IList<ExcelDiffSheetPair> sheetPairs,
            ExcelDiffOptions options,
            ExcelDiffOutputOptions output)
        {
            Baseline = baseline;
            Candidate = candidate;
            SheetPairs = sheetPairs ?? new List<ExcelDiffSheetPair>();
            Options = options ?? new ExcelDiffOptions();
            Output = output;
        }

        internal ExcelDiffWorkbookSource Baseline { get; private set; }
        internal ExcelDiffWorkbookSource Candidate { get; private set; }
        internal IList<ExcelDiffSheetPair> SheetPairs { get; private set; }
        internal ExcelDiffOptions Options { get; private set; }
        internal ExcelDiffOutputOptions Output { get; private set; }
    }

    internal sealed class ExcelDiffSheetExecutionResult
    {
        internal ExcelDiffSheetExecutionResult(ExcelDiffSheetPair pair, ExcelDiffResult result)
        {
            Pair = pair;
            Result = result;
        }

        internal ExcelDiffSheetPair Pair { get; private set; }
        internal ExcelDiffResult Result { get; private set; }
    }

    internal sealed class ExcelDiffExecutionResult
    {
        private ExcelDiffExecutionResult()
        {
            SheetResults = new List<ExcelDiffSheetExecutionResult>();
            Log = new ExcelDiffLog();
        }

        internal ExcelDiffOperationStatus Status { get; private set; }
        internal int DifferenceCount { get; private set; }
        internal int AddedRowCount { get; private set; }
        internal int DeletedRowCount { get; private set; }
        internal int ChangedRowCount { get; private set; }
        internal string OutputPath { get; private set; }
        internal IList<ExcelDiffSheetExecutionResult> SheetResults { get; private set; }
        internal ExcelDiffLog Log { get; private set; }
        internal Exception Error { get; private set; }
        internal ExcelDiffCancellationTrace CancellationTrace { get; private set; }

        internal static ExcelDiffExecutionResult Completed(string outputPath, ExcelDiffLog log, IList<ExcelDiffSheetExecutionResult> sheets, int differenceCount, int addedRows = 0, int deletedRows = 0, int changedRows = 0, ExcelDiffCancellationTrace cancellationTrace = null)
        {
            return Create(ExcelDiffOperationStatus.Completed, outputPath, log, sheets, differenceCount, null, addedRows, deletedRows, changedRows, cancellationTrace);
        }

        internal static ExcelDiffExecutionResult Cancelled(ExcelDiffLog log, IList<ExcelDiffSheetExecutionResult> sheets, int differenceCount, ExcelDiffCancellationTrace cancellationTrace = null)
        {
            return Create(ExcelDiffOperationStatus.Cancelled, null, log, sheets, differenceCount, null, 0, 0, 0, cancellationTrace);
        }

        internal static ExcelDiffExecutionResult Failed(ExcelDiffLog log, IList<ExcelDiffSheetExecutionResult> sheets, int differenceCount, Exception error, int addedRows = 0, int deletedRows = 0, int changedRows = 0, ExcelDiffCancellationTrace cancellationTrace = null)
        {
            return Create(ExcelDiffOperationStatus.Failed, null, log, sheets, differenceCount, error, addedRows, deletedRows, changedRows, cancellationTrace);
        }

        private static ExcelDiffExecutionResult Create(
            ExcelDiffOperationStatus status,
            string outputPath,
            ExcelDiffLog log,
            IList<ExcelDiffSheetExecutionResult> sheets,
            int differenceCount,
            Exception error,
            int addedRows = 0,
            int deletedRows = 0,
            int changedRows = 0,
            ExcelDiffCancellationTrace cancellationTrace = null)
        {
            ExcelDiffExecutionResult result = new ExcelDiffExecutionResult();
            result.Status = status;
            result.OutputPath = outputPath;
            result.Log = log ?? new ExcelDiffLog();
            result.SheetResults = sheets ?? new List<ExcelDiffSheetExecutionResult>();
            result.DifferenceCount = differenceCount;
            result.AddedRowCount = addedRows;
            result.DeletedRowCount = deletedRows;
            result.ChangedRowCount = changedRows;
            result.Error = error;
            result.CancellationTrace = cancellationTrace;
            return result;
        }
    }

    /// <summary>
    /// M2 业务服务。Excel COM 访问在调用线程完成；托管比较由 ExcelDiffEngine 执行。
    /// </summary>
    internal sealed class ExcelDiffService
    {
        private readonly Excel.Application application;
        private readonly ExcelDiffInteropPrototypeReader reader = new ExcelDiffInteropPrototypeReader();
        private readonly ExcelDiffEngine engine = new ExcelDiffEngine();
        private readonly ExcelDiffMarker marker = new ExcelDiffMarker();

        internal ExcelDiffService(Excel.Application application)
        {
            this.application = application ?? throw new ArgumentNullException("application");
        }

        internal ExcelDiffExecutionResult Compare(
            ExcelDiffRequest request,
            IExcelDiffProgress progress,
            CancellationToken cancellationToken)
        {
            ExcelDiffLog log = new ExcelDiffLog();
            IList<ExcelDiffSheetExecutionResult> sheetResults = new List<ExcelDiffSheetExecutionResult>();
            ExcelDiffWorkbookLease baselineLease = null;
            ExcelDiffWorkbookLease candidateLease = null;
            ExcelDiffWorkbookLease outputLease = null;
            string temporaryPath = null;
            int totalDifferences = 0;
            int totalAddedRows = 0;
            int totalDeletedRows = 0;
            int totalChangedRows = 0;
            ExcelDiffCancellationTrace cancellationTrace = progress == null || progress.CancellationTrace == null
                ? new ExcelDiffCancellationTrace(progress == null ? Guid.NewGuid() : progress.TaskId)
                : progress.CancellationTrace;
            try
            {
                ValidateRequest(request);
                Report(progress, ExcelDiffProgressStage.CheckingParameters, "正在检查工作簿和比较参数", string.Empty, 0, request.SheetPairs.Count, 0, null, 0, 0, 0, 0);
                log.Add("Excel 差异对比开始");

                string outputPath = NormalizeOutputPath(request.Output.OutputPath);
                ValidateOutputPath(outputPath, request);
                EnsureOutputExtension(outputPath, request.Candidate);
                CheckCancellation(cancellationToken, cancellationTrace);

                Report(progress, ExcelDiffProgressStage.InitializingExcel, "正在初始化 Excel 比较环境", string.Empty, 0, request.SheetPairs.Count, 0, null, 0, 0, 0, 0);
                Report(progress, ExcelDiffProgressStage.OpeningBaseline, "正在打开基准工作簿", string.Empty, 0, request.SheetPairs.Count, 0, null, 0, 0, 0, 0);
                baselineLease = OpenSource(request.Baseline, true, cancellationToken, cancellationTrace);
                CheckCancellation(cancellationToken, cancellationTrace);
                Report(progress, ExcelDiffProgressStage.OpeningCandidate, "正在打开待比较工作簿", string.Empty, 0, request.SheetPairs.Count, 0, null, 0, 0, 0, 0);
                candidateLease = OpenSource(request.Candidate, true, cancellationToken, cancellationTrace);
                CheckCancellation(cancellationToken, cancellationTrace);
                Report(progress, ExcelDiffProgressStage.LoadingSheets, "正在加载 Sheet 信息", string.Empty, 0, request.SheetPairs.Count, 0, null, 0, 0, 0, 0);
                ValidatePairs(baselineLease.Workbook, candidateLease.Workbook, request.SheetPairs, cancellationToken, cancellationTrace);

                Report(progress, ExcelDiffProgressStage.CopyingResult, "正在创建结果副本", string.Empty, 0, request.SheetPairs.Count, 0, null, 0, 0, 0, 0);
                temporaryPath = CreateTemporaryPath(outputPath);
                CopyCandidateToTemporary(candidateLease.Workbook, request.Candidate, temporaryPath, cancellationToken, cancellationTrace);
                outputLease = OpenWorkbook(temporaryPath, false, cancellationToken, cancellationTrace);
                CheckCancellation(cancellationToken, cancellationTrace);
                ValidateWritableSheets(outputLease.Workbook, request.SheetPairs, cancellationToken, cancellationTrace);

                for (int index = 0; index < request.SheetPairs.Count; index++)
                {
                    CheckCancellation(cancellationToken, cancellationTrace);
                    ExcelDiffSheetPair pair = request.SheetPairs[index];
                    Excel.Worksheet baselineSheet = null;
                    Excel.Worksheet candidateSheet = null;
                    Excel.Worksheet outputSheet = null;
                    try
                    {
                        Report(progress, ExcelDiffProgressStage.ReadingBaselineSheet, "正在读取基准 Sheet", pair.BaselineSheetName, index + 1, request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows);
                        baselineSheet = GetWorksheet(baselineLease.Workbook, pair.BaselineSheetName);
                        candidateSheet = GetWorksheet(candidateLease.Workbook, pair.CandidateSheetName);
                        outputSheet = GetWorksheet(outputLease.Workbook, pair.CandidateSheetName);
                        CheckCancellation(cancellationToken, cancellationTrace);
                        ExcelSheetData baselineData = reader.ReadWorksheet(baselineSheet, cancellationToken);
                        CheckCancellation(cancellationToken, cancellationTrace);
                        Report(progress, ExcelDiffProgressStage.ReadingCandidateSheet, "正在读取待比较 Sheet", pair.CandidateSheetName, index + 1, request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows);
                        ExcelSheetData candidateData = reader.ReadWorksheet(candidateSheet, cancellationToken);
                        CheckCancellation(cancellationToken, cancellationTrace);

                        long comparisonTotal = diffCandidateSize(baselineData, candidateData);
                        Report(progress, ExcelDiffProgressStage.Comparing, "正在比较单元格", pair.CandidateSheetName, index + 1, request.SheetPairs.Count, 0, comparisonTotal, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows);
                        ExcelDiffResult diff = engine.Compare(
                            baselineData,
                            candidateData,
                            cancellationToken,
                            processed => Report(progress, ExcelDiffProgressStage.ComparingCells, "正在比较单元格", pair.CandidateSheetName, index + 1, request.SheetPairs.Count, processed, comparisonTotal, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows));
                        CheckCancellation(cancellationToken, cancellationTrace);
                        totalDifferences += diff.Differences.Count;
                        totalAddedRows += diff.AddedRowCount;
                        totalDeletedRows += diff.DeletedRowCount;
                        totalChangedRows += diff.ChangedRowCount;
                        sheetResults.Add(new ExcelDiffSheetExecutionResult(pair, diff));
                        log.Add(string.Format(
                            "Sheet {0}: {1} 个差异",
                            index + 1,
                            diff.Differences.Count.ToString(CultureInfo.InvariantCulture)));

                        CheckCancellation(cancellationToken, cancellationTrace);
                        Report(progress, ExcelDiffProgressStage.InsertingStatusColumn, "正在插入行状态列", pair.CandidateSheetName, index + 1, request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows);
                        Report(progress, ExcelDiffProgressStage.WritingRowStatuses, "正在写入 ADD / DEL / CHG", pair.CandidateSheetName, index + 1, request.SheetPairs.Count, diff.Rows.Count, diff.Rows.Count, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows);
                        if (request.Options.AddBaselineComment)
                            Report(progress, ExcelDiffProgressStage.WritingComments, "正在写入原始值备注", pair.CandidateSheetName, index + 1, request.SheetPairs.Count, diff.Differences.Count, diff.Differences.Count, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows);
                        marker.Apply(outputSheet, diff, request.Options, cancellationToken);
                        CheckCancellation(cancellationToken, cancellationTrace);
                    }
                    finally
                    {
                        ComHelper.Release(outputSheet);
                        ComHelper.Release(candidateSheet);
                        ComHelper.Release(baselineSheet);
                    }

                    Report(progress, ExcelDiffProgressStage.Marking, "已完成当前 Sheet 的差异标记", pair.CandidateSheetName, index + 1, request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows);
                }

                CheckCancellation(cancellationToken, cancellationTrace);
                Report(progress, ExcelDiffProgressStage.Saving, "正在保存结果工作簿", string.Empty, request.SheetPairs.Count, request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows);
                SaveAndClose(outputLease, cancellationToken, cancellationTrace);
                outputLease = null;
                CheckCancellation(cancellationToken, cancellationTrace);
                if (request.Options.ValidateOutputAfterSave)
                {
                    ValidateSavedWorkbook(temporaryPath, cancellationToken, cancellationTrace);
                }

                CheckCancellation(cancellationToken, cancellationTrace);
                PublishTemporaryFile(temporaryPath, outputPath, request.Options.OverwriteExistingOutput);
                temporaryPath = null;
                log.Add("Excel 差异对比完成");
                Report(progress, ExcelDiffProgressStage.Completed, "比较完成", string.Empty, request.SheetPairs.Count, request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows, ExcelDiffProgressStatus.Completed);
                return ExcelDiffExecutionResult.Completed(outputPath, log, sheetResults, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows, cancellationTrace);
            }
            catch (OperationCanceledException)
            {
                log.Add("Excel 差异对比已取消");
                cancellationTrace.RecordOnce(ExcelDiffCancellationEventKind.CancelObserved, cancellationTrace.CurrentStage, "工作线程感知取消");
                cancellationTrace.RecordOnce(ExcelDiffCancellationEventKind.BusinessStopped, cancellationTrace.CurrentStage, "已停止后续比较、标记和发布");
                Report(progress, ExcelDiffProgressStage.Cancelled, "比较已取消，等待资源清理", string.Empty, sheetResults.Count, request == null ? 0 : request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows, ExcelDiffProgressStatus.Cancelled);
                return ExcelDiffExecutionResult.Cancelled(log, sheetResults, totalDifferences, cancellationTrace);
            }
            catch (Exception exception)
            {
                log.Add("Excel 差异对比失败：" + exception.GetType().FullName);
                Report(progress, ExcelDiffProgressStage.Failed, "比较失败", string.Empty, sheetResults.Count, request == null ? 0 : request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows, ExcelDiffProgressStatus.Failed);
                return ExcelDiffExecutionResult.Failed(log, sheetResults, totalDifferences, exception, totalAddedRows, totalDeletedRows, totalChangedRows, cancellationTrace);
            }
            finally
            {
                cancellationTrace.RecordOnce(ExcelDiffCancellationEventKind.CleanupStarted, cancellationTrace.CurrentStage, "正在清理临时文件和 Excel 资源");
                if (cancellationTrace.IsCancellationRequested)
                    Report(progress, ExcelDiffProgressStage.CleaningUp, "比较已停止，正在清理临时文件和 Excel 资源", string.Empty, sheetResults.Count, request == null ? 0 : request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows, ExcelDiffProgressStatus.Cancelling);
                CloseLeaseSafely(outputLease, log);
                CloseLeaseSafely(candidateLease, log);
                CloseLeaseSafely(baselineLease, log);
                try { DeleteTemporaryFile(temporaryPath); }
                catch (Exception cleanupException) { log.Add("临时文件清理失败：" + cleanupException.GetType().FullName); }
                cancellationTrace.RecordOnce(ExcelDiffCancellationEventKind.CleanupCompleted, cancellationTrace.CurrentStage, "资源清理完成");
                if (cancellationTrace.IsCancellationRequested)
                    Report(progress, ExcelDiffProgressStage.CleaningUp, "清理完成，等待工作线程退出", string.Empty, sheetResults.Count, request == null ? 0 : request.SheetPairs.Count, 0, null, totalDifferences, totalAddedRows, totalDeletedRows, totalChangedRows, ExcelDiffProgressStatus.Cancelled);
            }
        }

        private static void CloseLeaseSafely(ExcelDiffWorkbookLease lease, ExcelDiffLog log)
        {
            if (lease == null) return;
            try { lease.Close(); }
            catch (Exception exception) { log.Add("Excel 资源清理失败：" + exception.GetType().FullName); }
        }

        private void ValidateRequest(ExcelDiffRequest request)
        {
            if (request == null) throw new ArgumentNullException("request");
            if (request.Baseline == null) throw new ArgumentException("缺少基准工作簿", "request");
            if (request.Candidate == null) throw new ArgumentException("缺少待比较工作簿", "request");
            if (request.Output == null || string.IsNullOrWhiteSpace(request.Output.OutputPath)) throw new ArgumentException("缺少输出路径", "request");
            if (request.SheetPairs == null || request.SheetPairs.Count == 0) throw new ArgumentException("至少需要一个 Sheet 映射", "request");
            foreach (ExcelDiffSheetPair pair in request.SheetPairs)
            {
                if (pair == null || string.IsNullOrWhiteSpace(pair.BaselineSheetName) || string.IsNullOrWhiteSpace(pair.CandidateSheetName))
                {
                    throw new ArgumentException("Sheet 映射不能为空", "request");
                }
            }
        }

        private void ValidateOutputPath(string outputPath, ExcelDiffRequest request)
        {
            string directory = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(directory)) throw new ArgumentException("输出路径缺少目录");
            Directory.CreateDirectory(directory);
            if (File.Exists(outputPath) && !request.Options.OverwriteExistingOutput)
            {
                throw new IOException("输出文件已存在，未启用覆盖选项");
            }

            string[] sourcePaths = { GetSourcePath(request.Baseline), GetSourcePath(request.Candidate) };
            foreach (string sourcePath in sourcePaths)
            {
                if (!string.IsNullOrEmpty(sourcePath) && PathsEqual(sourcePath, outputPath))
                {
                    throw new IOException("输出路径不能与源工作簿相同");
                }
            }
        }

        private void EnsureOutputExtension(string outputPath, ExcelDiffWorkbookSource candidate)
        {
            string candidatePath = GetSourcePath(candidate);
            if (string.IsNullOrEmpty(candidatePath)) return;
            string candidateExtension = Path.GetExtension(candidatePath);
            string outputExtension = Path.GetExtension(outputPath);
            if (IsMacroExtension(candidateExtension) && !string.Equals(candidateExtension, outputExtension, StringComparison.OrdinalIgnoreCase))
            {
                throw new IOException("含宏工作簿必须使用相同扩展名输出，以避免丢失 VBA 内容");
            }
        }

        private void ValidatePairs(Excel.Workbook baseline, Excel.Workbook candidate, IList<ExcelDiffSheetPair> pairs, CancellationToken cancellationToken, ExcelDiffCancellationTrace cancellationTrace)
        {
            HashSet<string> baselineNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> candidateNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ExcelDiffSheetPair pair in pairs)
            {
                CheckCancellation(cancellationToken, cancellationTrace);
                Excel.Worksheet baselineSheet = null;
                Excel.Worksheet candidateSheet = null;
                try
                {
                    baselineSheet = GetWorksheet(baseline, pair.BaselineSheetName);
                    candidateSheet = GetWorksheet(candidate, pair.CandidateSheetName);
                    if (!baselineNames.Add(pair.BaselineSheetName) || !candidateNames.Add(pair.CandidateSheetName))
                    {
                        throw new ArgumentException("Sheet 映射必须是一对一且不能重复");
                    }
                }
                finally
                {
                    ComHelper.Release(candidateSheet);
                    ComHelper.Release(baselineSheet);
                }
            }
        }

        private void ValidateWritableSheets(Excel.Workbook workbook, IList<ExcelDiffSheetPair> pairs, CancellationToken cancellationToken, ExcelDiffCancellationTrace cancellationTrace)
        {
            foreach (ExcelDiffSheetPair pair in pairs)
            {
                CheckCancellation(cancellationToken, cancellationTrace);
                Excel.Worksheet sheet = null;
                try
                {
                    sheet = GetWorksheet(workbook, pair.CandidateSheetName);
                    if (sheet.ProtectContents)
                    {
                        throw new InvalidOperationException("结果 Sheet 受保护，无法标记差异");
                    }
                }
                finally
                {
                    ComHelper.Release(sheet);
                }
            }
        }

        private ExcelDiffWorkbookLease OpenSource(ExcelDiffWorkbookSource source, bool readOnly, CancellationToken cancellationToken, ExcelDiffCancellationTrace cancellationTrace)
        {
            CheckCancellation(cancellationToken, cancellationTrace);
            if (source.Workbook != null)
            {
                return new ExcelDiffWorkbookLease(source.Workbook, false);
            }
            if (string.IsNullOrEmpty(source.FilePath)) throw new ArgumentException("工作簿来源为空");
            string path = Path.GetFullPath(source.FilePath);
            if (!File.Exists(path)) throw new FileNotFoundException("工作簿不存在", path);
            return OpenWorkbook(path, readOnly, cancellationToken, cancellationTrace);
        }

        private ExcelDiffWorkbookLease OpenWorkbook(string path, bool readOnly, CancellationToken cancellationToken, ExcelDiffCancellationTrace cancellationTrace)
        {
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            try
            {
                CheckCancellation(cancellationToken, cancellationTrace);
                workbooks = application.Workbooks;
                workbook = workbooks.Open(
                    path,
                    UpdateLinks: 0,
                    ReadOnly: readOnly,
                    AddToMru: false,
                    IgnoreReadOnlyRecommended: true);
                CheckCancellation(cancellationToken, cancellationTrace);
                return new ExcelDiffWorkbookLease(workbook, true);
            }
            catch
            {
                ComHelper.Release(workbook);
                throw;
            }
            finally
            {
                ComHelper.Release(workbooks);
            }
        }

        private void CopyCandidateToTemporary(Excel.Workbook candidateWorkbook, ExcelDiffWorkbookSource source, string temporaryPath, CancellationToken cancellationToken, ExcelDiffCancellationTrace cancellationTrace)
        {
            CheckCancellation(cancellationToken, cancellationTrace);
            string sourcePath = GetSourcePath(source);
            if (!string.IsNullOrEmpty(sourcePath))
            {
                File.Copy(sourcePath, temporaryPath, false);
            }
            else
            {
                candidateWorkbook.SaveCopyAs(temporaryPath);
            }
            CheckCancellation(cancellationToken, cancellationTrace);
        }

        private void SaveAndClose(ExcelDiffWorkbookLease outputLease, CancellationToken cancellationToken, ExcelDiffCancellationTrace cancellationTrace)
        {
            if (outputLease == null) return;
            CheckCancellation(cancellationToken, cancellationTrace);
            outputLease.Workbook.Save();
            CheckCancellation(cancellationToken, cancellationTrace);
            outputLease.Close();
        }

        private void ValidateSavedWorkbook(string path, CancellationToken cancellationToken, ExcelDiffCancellationTrace cancellationTrace)
        {
            ExcelDiffWorkbookLease validation = null;
            try
            {
                CheckCancellation(cancellationToken, cancellationTrace);
                validation = OpenWorkbook(path, true, cancellationToken, cancellationTrace);
            }
            finally
            {
                if (validation != null) validation.Close();
            }
        }

        private static void PublishTemporaryFile(string temporaryPath, string outputPath, bool overwrite)
        {
            if (overwrite && File.Exists(outputPath))
            {
                File.Replace(temporaryPath, outputPath, null);
            }
            else
            {
                File.Move(temporaryPath, outputPath);
            }
        }

        private static string CreateTemporaryPath(string outputPath)
        {
            string extension = Path.GetExtension(outputPath);
            string directory = Path.GetDirectoryName(outputPath);
            string fileName = Path.GetFileNameWithoutExtension(outputPath);
            return Path.Combine(directory, "." + fileName + "." + Guid.NewGuid().ToString("N") + extension);
        }

        private static string NormalizeOutputPath(string path)
        {
            return Path.GetFullPath(path.Trim());
        }

        private static string GetSourcePath(ExcelDiffWorkbookSource source)
        {
            if (source == null) return null;
            if (!string.IsNullOrEmpty(source.FilePath)) return Path.GetFullPath(source.FilePath);
            if (source.Workbook == null) return null;
            return source.Workbook.Path;
        }

        private static bool PathsEqual(string first, string second)
        {
            return string.Equals(Path.GetFullPath(first), Path.GetFullPath(second), StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsMacroExtension(string extension)
        {
            return string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".xltm", StringComparison.OrdinalIgnoreCase);
        }

        private static Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string name)
        {
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                Excel.Worksheet worksheet = sheets[name] as Excel.Worksheet;
                if (worksheet == null) throw new ArgumentException("Sheet 不存在：" + name);
                return worksheet;
            }
            finally
            {
                ComHelper.Release(sheets);
            }
        }

        private static void DeleteTemporaryFile(string path)
        {
            if (string.IsNullOrEmpty(path)) return;
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch (IOException)
            {
                // 清理失败不覆盖主异常；结果状态仍为失败，调用方可根据日志处理。
            }
            catch (UnauthorizedAccessException)
            {
                // 同上。
            }
        }

        private static long diffCandidateSize(ExcelSheetData baseline, ExcelSheetData candidate)
        {
            long rows = Math.Max(baseline == null ? 0 : baseline.Rows.Count, candidate == null ? 0 : candidate.Rows.Count);
            long columns = 0;
            for (int row = 0; row < rows; row++)
            {
                int baselineColumns = baseline != null && row < baseline.Rows.Count && baseline.Rows[row] != null ? baseline.Rows[row].Count : 0;
                int candidateColumns = candidate != null && row < candidate.Rows.Count && candidate.Rows[row] != null ? candidate.Rows[row].Count : 0;
                columns = Math.Max(columns, Math.Max(baselineColumns, candidateColumns));
            }
            return rows * columns;
        }

        private static void CheckCancellation(CancellationToken cancellationToken, ExcelDiffCancellationTrace cancellationTrace)
        {
            if (!cancellationToken.IsCancellationRequested) return;
            if (!cancellationTrace.IsCancellationRequested)
                cancellationTrace.RecordCancelRequested(cancellationTrace.CurrentStage, "服务检测到取消请求");
            cancellationTrace.RecordOnce(
                ExcelDiffCancellationEventKind.CancelObserved,
                cancellationTrace.CurrentStage,
                "工作线程在安全检查点感知取消");
            throw new OperationCanceledException(cancellationToken);
        }

        private static void Report(
            IExcelDiffProgress progress,
            ExcelDiffProgressStage stage,
            string operation,
            string sheetName,
            int completed,
            int total,
            long processed,
            long? workTotal,
            int differences,
            int addedRows,
            int deletedRows,
            int changedRows,
            ExcelDiffProgressStatus status = ExcelDiffProgressStatus.Running)
        {
            if (progress != null)
            {
                if (progress.CancellationTrace != null) progress.CancellationTrace.SetCurrentStage(stage, operation);
                progress.Report(new ExcelDiffProgressInfo(stage, operation, sheetName, completed, total, processed, workTotal, differences, addedRows, deletedRows, changedRows, status));
            }
        }

        private sealed class ExcelDiffWorkbookLease
        {
            internal ExcelDiffWorkbookLease(Excel.Workbook workbook, bool ownsWorkbook)
            {
                Workbook = workbook;
                OwnsWorkbook = ownsWorkbook;
            }

            internal Excel.Workbook Workbook { get; private set; }
            private bool OwnsWorkbook { get; set; }

            internal void Close()
            {
                if (!OwnsWorkbook || Workbook == null) return;
                try
                {
                    Workbook.Close(false);
                }
                finally
                {
                    ComHelper.Release(Workbook);
                    Workbook = null;
                    OwnsWorkbook = false;
                }
            }
        }

        private sealed class ExcelDiffMarker
        {
            internal void Apply(Excel.Worksheet worksheet, ExcelDiffResult result, ExcelDiffOptions options, CancellationToken cancellationToken)
            {
                if (worksheet == null) throw new ArgumentNullException("worksheet");
                if (result == null) throw new ArgumentNullException("result");
                if (worksheet.ProtectContents) throw new InvalidOperationException("结果 Sheet 受保护，无法标记差异");
                cancellationToken.ThrowIfCancellationRequested();

                Excel.ListObjects tables = null;
                Excel.Range firstColumn = null;
                try
                {
                    tables = worksheet.ListObjects;
                    if (tables != null && tables.Count > 0)
                        throw new InvalidOperationException("结果 Sheet 包含 Excel Table，无法安全插入行状态列；请先转换为普通区域后重试");

                    firstColumn = worksheet.Columns[1] as Excel.Range;
                    firstColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                    firstColumn.ColumnWidth = 10;
                    cancellationToken.ThrowIfCancellationRequested();
                }
                finally
                {
                    ComHelper.Release(firstColumn);
                    ComHelper.Release(tables);
                }

                foreach (ExcelDiffRow row in result.Rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    WriteRowStatus(worksheet, row);
                }

                int differenceIndex = 0;
                foreach (ExcelDiffCell difference in result.Differences)
                {
                    if ((differenceIndex++ & 15) == 0) cancellationToken.ThrowIfCancellationRequested();
                    ApplyOne(worksheet, difference, options);
                }
            }

            private static void WriteRowStatus(Excel.Worksheet worksheet, ExcelDiffRow row)
            {
                Excel.Range cell = null;
                Excel.Interior interior = null;
                try
                {
                    cell = worksheet.Cells[row.RowNumber, 1];
                    cell.Value2 = row.StatusText;
                    cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    interior = cell.Interior;
                    switch (row.Status)
                    {
                        case ExcelDiffRowStatus.Added: interior.Color = ColorTranslator.ToOle(Color.LightGreen); break;
                        case ExcelDiffRowStatus.Deleted: interior.Color = ColorTranslator.ToOle(Color.LightSalmon); break;
                        case ExcelDiffRowStatus.Changed: interior.Color = ColorTranslator.ToOle(Color.Khaki); break;
                    }
                }
                finally
                {
                    ComHelper.Release(interior);
                    ComHelper.Release(cell);
                }
            }

            private static void ApplyOne(Excel.Worksheet worksheet, ExcelDiffCell difference, ExcelDiffOptions options)
            {
                Excel.Range cell = null;
                Excel.Range worksheetCells = null;
                Excel.Range target = null;
                Excel.Range mergeArea = null;
                Excel.Range mergeCells = null;
                Excel.Range topLeft = null;
                Excel.Interior interior = null;
                Excel.Comment comment = null;
                try
                {
                    worksheetCells = worksheet.Cells;
                    cell = worksheetCells[difference.Coordinate.Row, difference.Coordinate.Column + 1];
                    target = cell;
                    bool isMerged = cell.MergeCells is bool && (bool)cell.MergeCells;
                    if (isMerged)
                    {
                        mergeArea = cell.MergeArea;
                        target = mergeArea;
                        mergeCells = mergeArea.Cells;
                        topLeft = mergeCells[1, 1] as Excel.Range;
                    }
                    else
                    {
                        topLeft = cell;
                    }

                    if (options.HighlightDifferences)
                    {
                        interior = target.Interior;
                        interior.Color = ColorTranslator.ToOle(options.HighlightColor);
                    }

                    if (options.AddBaselineComment && !string.IsNullOrEmpty(difference.BaselineValue.ComparisonText))
                    {
                        string text = "旧数据: " + Environment.NewLine + difference.BaselineValue.ComparisonText;
                        comment = topLeft.Comment;
                        if (comment == null)
                        {
                            comment = topLeft.AddComment(text);
                        }
                        else
                        {
                            string existing = comment.Text();
                            comment.Text(existing + Environment.NewLine + "[ExcelDiff]" + Environment.NewLine + text);
                        }
                    }
                }
                finally
                {
                    ComHelper.Release(comment);
                    ComHelper.Release(interior);
                    if (!ReferenceEquals(topLeft, cell)) ComHelper.Release(topLeft);
                    ComHelper.Release(mergeCells);
                    if (!ReferenceEquals(target, cell)) ComHelper.Release(target);
                    ComHelper.Release(mergeArea);
                    ComHelper.Release(worksheetCells);
                    ComHelper.Release(cell);
                }
            }
        }
    }

}
