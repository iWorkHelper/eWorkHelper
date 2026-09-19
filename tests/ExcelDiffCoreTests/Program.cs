using System;
using System.Collections.Generic;
using System.Diagnostics;
using eWorkhelper;

internal static class Program
{
    private static int passed;

    private static void Main()
    {
        Run("相同 Sheet 无差异", SameSheetHasNoDifferences);
        Run("单点变化", SingleChangedCell);
        Run("多点变化和类型", MultipleChangesAndKinds);
        Run("基准侧有值候选侧为空", RemovedValue);
        Run("候选侧新增值", AddedValue);
        Run("缺失和空字符串按参考语义相等", MissingAndEmptyAreEqual);
        Run("零值和文本零不同", NumericAndTextZeroAreDifferent);
        Run("不同边界", DifferentBounds);
        Run("空 Sheet", EmptySheets);
        Run("中间插行保持坐标比较", InsertedRowIsCoordinateBased);
        Run("行状态聚合", RowStatusAggregation);
        Run("部分单元格清空为修改行", PartialRowClearIsChanged);
        Run("独立监控计时不依赖工作线程回调", IndependentMonitorTicks);
        Run("托管比较按批次响应取消", ManagedCompareCancellation);
        Run("取消后普通进度不覆盖取消状态", CancellationStateWinsOverProgress);
        Run("合并元数据不影响核心值比较", MergeMetadataDoesNotEnterCore);
        Run("参数校验", NullArguments);
        Run("规模基准 100x20/1000x20/10000x20", PerformanceSmoke);

        Console.WriteLine("PASS: {0}", passed);
    }

    private static void SameSheetHasNoDifferences()
    {
        ExcelDiffResult result = Compare(
            Rows(new[] { "A", "1" }, new[] { "B", "2" }),
            Rows(new[] { "A", "1" }, new[] { "B", "2" }));
        AssertEqual(0, result.Differences.Count, "difference count");
        AssertEqual(2, result.ComparedRowCount, "row boundary");
        AssertEqual(2, result.ComparedColumnCount, "column boundary");
    }

    private static void SingleChangedCell()
    {
        ExcelDiffResult result = Compare(Rows(new[] { "A", "old" }), Rows(new[] { "A", "new" }));
        AssertEqual(1, result.Differences.Count, "difference count");
        AssertEqual("R1C2", result.Differences[0].Coordinate.ToString(), "coordinate");
        AssertEqual("old", result.Differences[0].BaselineValue.ComparisonText, "baseline");
        AssertEqual("new", result.Differences[0].CandidateValue.ComparisonText, "candidate");
        AssertEqual(ExcelDiffKind.Changed, result.Differences[0].Kind, "kind");
    }

    private static void MultipleChangesAndKinds()
    {
        ExcelDiffResult result = Compare(
            Rows(new[] { "a", "1", "old" }),
            Rows(new[] { "b", "1", "new" }));
        AssertEqual(2, result.Differences.Count, "difference count");
        AssertEqual(ExcelDiffKind.Changed, result.Differences[0].Kind, "first kind");
        AssertEqual(ExcelDiffKind.Changed, result.Differences[1].Kind, "second kind");
    }

    private static void RemovedValue()
    {
        ExcelDiffResult result = Compare(Rows(new[] { "old" }), Rows(new[] { "" }));
        AssertEqual(1, result.Differences.Count, "difference count");
        AssertEqual(ExcelDiffKind.Removed, result.Differences[0].Kind, "kind");
    }

    private static void AddedValue()
    {
        ExcelDiffResult result = Compare(Rows(new[] { "" }), Rows(new[] { "new" }));
        AssertEqual(1, result.Differences.Count, "difference count");
        AssertEqual(ExcelDiffKind.Added, result.Differences[0].Kind, "kind");
    }

    private static void MissingAndEmptyAreEqual()
    {
        ExcelSheetData baseline = ExcelSheetData.FromComparisonRows("old", new string[0][]);
        ExcelSheetData candidate = ExcelSheetData.FromComparisonRows("new", new[] { new[] { "" } });
        ExcelDiffResult result = new ExcelDiffEngine().Compare(baseline, candidate);
        AssertEqual(1, result.ComparedRowCount, "row boundary");
        AssertEqual(0, result.Differences.Count, "missing/empty comparison");
    }

    private static void NumericAndTextZeroAreDifferent()
    {
        ExcelSheetData baseline = new ExcelSheetData("old", new List<IList<ExcelCellValue>>
        {
            new List<ExcelCellValue> { new ExcelCellValue(true, "0", ExcelCellSourceKind.Number, 0d) }
        });
        ExcelSheetData candidate = ExcelSheetData.FromComparisonRows("new", new[] { new[] { "0" } });
        ExcelDiffResult result = new ExcelDiffEngine().Compare(baseline, candidate);
        AssertEqual(0, result.Differences.Count, "reference string semantics");
    }

    private static void DifferentBounds()
    {
        ExcelDiffResult result = Compare(Rows(new[] { "a" }), Rows(new[] { "a", "b" }, new[] { "c", "d" }));
        AssertEqual(3, result.Differences.Count, "expanded bounds differences");
        AssertEqual(2, result.ComparedRowCount, "row boundary");
        AssertEqual(2, result.ComparedColumnCount, "column boundary");
    }

    private static void EmptySheets()
    {
        ExcelDiffResult result = Compare(Rows(), Rows());
        AssertEqual(0, result.ComparedRowCount, "row boundary");
        AssertEqual(0, result.ComparedColumnCount, "column boundary");
        AssertEqual(0, result.Differences.Count, "difference count");
    }

    private static void InsertedRowIsCoordinateBased()
    {
        ExcelDiffResult result = Compare(
            Rows(new[] { "id", "one" }, new[] { "id2", "two" }),
            Rows(new[] { "id", "inserted" }, new[] { "new", "row" }, new[] { "id2", "two" }));
        AssertEqual(5, result.Differences.Count, "coordinate comparison count");
    }

    private static void RowStatusAggregation()
    {
        ExcelDiffResult result = Compare(
            Rows(new[] { "same" }, new[] { "deleted" }, new string[0], new[] { "old" }),
            Rows(new[] { "same" }, new string[0], new[] { "added" }, new[] { "new" }));
        AssertEqual(3, result.Rows.Count, "non-unchanged row count");
        AssertEqual(ExcelDiffRowStatus.Deleted, result.Rows[0].Status, "deleted row");
        AssertEqual(ExcelDiffRowStatus.Added, result.Rows[1].Status, "added row");
        AssertEqual(ExcelDiffRowStatus.Changed, result.Rows[2].Status, "changed row");
        AssertEqual(1, result.AddedRowCount, "added count");
        AssertEqual(1, result.DeletedRowCount, "deleted count");
        AssertEqual(1, result.ChangedRowCount, "changed count");
    }

    private static void PartialRowClearIsChanged()
    {
        ExcelDiffResult result = Compare(
            Rows(new[] { "id", "name", "amount" }),
            Rows(new[] { "id", "name", "" }));
        AssertEqual(1, result.ChangedRowCount, "partial clear is CHG");
        AssertEqual(0, result.DeletedRowCount, "partial clear is not DEL");
        AssertEqual("CHG", result.Rows[0].StatusText, "status text");
    }

    private static void IndependentMonitorTicks()
    {
        int callbacks = 0;
        ExcelDiffProgressSnapshot latest = null;
        using (ExcelDiffProgressMonitor monitor = new ExcelDiffProgressMonitor(snapshot => { callbacks++; latest = snapshot; }))
        {
            monitor.Start();
            System.Threading.Thread.Sleep(1200);
            if (callbacks == 0) throw new InvalidOperationException("monitor callback without worker report");
            if (latest == null || latest.Elapsed < TimeSpan.FromSeconds(1)) throw new InvalidOperationException("monotonic elapsed time");
        }
    }

    private static void ManagedCompareCancellation()
    {
        ExcelSheetData baseline = CreateSheet(20000, 20, false);
        ExcelSheetData candidate = CreateSheet(20000, 20, true);
        System.Threading.CancellationTokenSource source = new System.Threading.CancellationTokenSource();
        bool cancelled = false;
        Stopwatch stopwatch = Stopwatch.StartNew();
        try
        {
            new ExcelDiffEngine().Compare(baseline, candidate, source.Token, processed =>
            {
                if (processed >= 256) source.Cancel();
            });
        }
        catch (OperationCanceledException)
        {
            cancelled = true;
            stopwatch.Stop();
            Console.WriteLine("CANCEL_PERF managed_compare_cancel_ms={0}", stopwatch.Elapsed.TotalMilliseconds);
        }
        if (!cancelled) throw new InvalidOperationException("managed compare did not observe cancellation");
    }

    private static void CancellationStateWinsOverProgress()
    {
        ExcelDiffProgressSnapshot latest = null;
        using (ExcelDiffProgressMonitor monitor = new ExcelDiffProgressMonitor(snapshot => latest = snapshot))
        {
            monitor.Start();
            monitor.Report(new ExcelDiffProgressInfo(ExcelDiffProgressStage.Comparing, "正在比较", "Data", 1, 1, 1, 10, 0, 0, 0, 0, ExcelDiffProgressStatus.Running));
            monitor.RequestCancellation();
            monitor.Report(new ExcelDiffProgressInfo(ExcelDiffProgressStage.WritingComments, "正在写备注", "Data", 1, 1, 2, 10, 1, 0, 0, 1, ExcelDiffProgressStatus.Running));
            if (latest == null || latest.Status != ExcelDiffProgressStatus.Cancelling)
                throw new InvalidOperationException("ordinary progress overwrote cancelling state");
            if (monitor.CancellationTrace.Snapshot().Count != 1)
                throw new InvalidOperationException("unexpected cancellation trace count");
        }
    }

    private static void MergeMetadataDoesNotEnterCore()
    {
        ExcelDiffResult result = Compare(Rows(new[] { "merged" }), Rows(new[] { "merged" }));
        AssertEqual(0, result.Differences.Count, "merge metadata is adapter concern");
    }

    private static void NullArguments()
    {
        AssertThrows<ArgumentNullException>(() => new ExcelDiffEngine().Compare(null, ExcelSheetData.FromComparisonRows("x")), "baseline null");
        AssertThrows<ArgumentNullException>(() => new ExcelDiffEngine().Compare(ExcelSheetData.FromComparisonRows("x"), null), "candidate null");
    }

    private static void PerformanceSmoke()
    {
        int[] rowCounts = { 100, 1000, 10000 };
        foreach (int rowCount in rowCounts)
        {
            ExcelSheetData baseline = CreateSheet(rowCount, 20, false);
            ExcelSheetData candidate = CreateSheet(rowCount, 20, rowCount == 10000);
            Stopwatch stopwatch = Stopwatch.StartNew();
            ExcelDiffResult result = new ExcelDiffEngine().Compare(baseline, candidate);
            stopwatch.Stop();
            AssertEqual(rowCount == 10000 ? 1 : 0, result.Differences.Count, "performance diff count");
            Console.WriteLine("PERF rows={0}, cols=20, compare_ms={1}, managed_bytes={2}", rowCount, stopwatch.Elapsed.TotalMilliseconds, GC.GetTotalMemory(false));
        }
    }

    private static ExcelSheetData CreateSheet(int rows, int columns, bool changeLastCell)
    {
        List<IList<ExcelCellValue>> values = new List<IList<ExcelCellValue>>();
        for (int row = 0; row < rows; row++)
        {
            List<ExcelCellValue> current = new List<ExcelCellValue>();
            for (int column = 0; column < columns; column++)
            {
                string value = row.ToString() + ":" + column.ToString();
                if (changeLastCell && row == rows - 1 && column == columns - 1)
                {
                    value = "changed";
                }
                current.Add(ExcelCellValue.FromText(value));
            }
            values.Add(current);
        }
        return new ExcelSheetData("Sheet1", values);
    }

    private static ExcelDiffResult Compare(ExcelSheetData baseline, ExcelSheetData candidate)
    {
        return new ExcelDiffEngine().Compare(baseline, candidate);
    }

    private static ExcelSheetData Rows(params string[][] rows)
    {
        return ExcelSheetData.FromComparisonRows("Sheet1", rows);
    }

    private static void Run(string name, Action test)
    {
        try
        {
            test();
            passed++;
            Console.WriteLine("PASS: " + name);
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine("FAIL: " + name + " - " + exception.Message);
            Environment.Exit(1);
        }
    }

    private static void AssertEqual<T>(T expected, T actual, string message)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
        {
            throw new InvalidOperationException(string.Format("{0}: expected {1}, actual {2}", message, expected, actual));
        }
    }

    private static void AssertThrows<T>(Action action, string message) where T : Exception
    {
        try
        {
            action();
        }
        catch (T)
        {
            return;
        }
        throw new InvalidOperationException(message + ": expected " + typeof(T).Name);
    }
}
