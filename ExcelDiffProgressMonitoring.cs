using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace eWorkhelper
{
    internal enum ExcelDiffProgressStage
    {
        Starting, CheckingParameters, InitializingExcel, OpeningBaseline, OpeningCandidate, LoadingSheets,
        ReadingBaselineSheet, ReadingCandidateSheet, ComparingCells, AggregatingRows, CopyingResult,
        InsertingStatusColumn, WritingRowStatuses, Opening, Reading, Comparing, Marking, WritingComments,
        Saving, CleaningUp, Completed, Cancelled, Failed
    }

    internal enum ExcelDiffProgressStatus { Running, Cancelling, Completed, Cancelled, Failed }

    internal sealed class ExcelDiffProgressInfo
    {
        internal ExcelDiffProgressInfo(ExcelDiffProgressStage stage, int completedSheets, int totalSheets, int differences)
            : this(stage, string.Empty, string.Empty, completedSheets, totalSheets, 0, null, differences, 0, 0, 0, ExcelDiffProgressStatus.Running) { }

        internal ExcelDiffProgressInfo(ExcelDiffProgressStage stage, string operation, string sheetName, int completedSheets, int totalSheets, long processed, long? total, int differences, int addedRows, int deletedRows, int changedRows, ExcelDiffProgressStatus status)
        {
            Stage = stage; Operation = operation ?? string.Empty; SheetName = sheetName ?? string.Empty;
            CompletedSheets = completedSheets; TotalSheets = totalSheets; Processed = processed; Total = total;
            Differences = differences; AddedRows = addedRows; DeletedRows = deletedRows; ChangedRows = changedRows; Status = status;
        }

        internal ExcelDiffProgressStage Stage { get; private set; }
        internal int CompletedSheets { get; private set; }
        internal int TotalSheets { get; private set; }
        internal int Differences { get; private set; }
        internal string Operation { get; private set; }
        internal string SheetName { get; private set; }
        internal long Processed { get; private set; }
        internal long? Total { get; private set; }
        internal int AddedRows { get; private set; }
        internal int DeletedRows { get; private set; }
        internal int ChangedRows { get; private set; }
        internal ExcelDiffProgressStatus Status { get; private set; }
    }

    internal enum ExcelDiffCancellationEventKind
    {
        CancelRequested,
        CancelObserved,
        BusinessStopped,
        CleanupStarted,
        CleanupCompleted,
        WorkerExited
    }

    internal sealed class ExcelDiffCancellationEvent
    {
        internal ExcelDiffCancellationEvent(ExcelDiffCancellationEventKind kind, Guid taskId, ExcelDiffProgressStage stage, string operation, DateTime utc, TimeSpan elapsed, TimeSpan? sinceCancel, int threadId)
        {
            Kind = kind;
            TaskId = taskId;
            Stage = stage;
            Operation = operation ?? string.Empty;
            Utc = utc;
            Elapsed = elapsed;
            SinceCancel = sinceCancel;
            ThreadId = threadId;
        }

        internal ExcelDiffCancellationEventKind Kind { get; private set; }
        internal Guid TaskId { get; private set; }
        internal ExcelDiffProgressStage Stage { get; private set; }
        internal string Operation { get; private set; }
        internal DateTime Utc { get; private set; }
        internal TimeSpan Elapsed { get; private set; }
        internal TimeSpan? SinceCancel { get; private set; }
        internal int ThreadId { get; private set; }
    }

    internal sealed class ExcelDiffCancellationTrace
    {
        private readonly object sync = new object();
        private readonly Stopwatch stopwatch = Stopwatch.StartNew();
        private readonly Guid taskId;
        private readonly IList<ExcelDiffCancellationEvent> events = new List<ExcelDiffCancellationEvent>();
        private TimeSpan? cancelElapsed;
        private ExcelDiffProgressStage currentStage = ExcelDiffProgressStage.Starting;
        private string currentOperation = string.Empty;

        internal ExcelDiffCancellationTrace(Guid taskId)
        {
            this.taskId = taskId;
        }

        internal Guid TaskId { get { return taskId; } }
        internal bool IsCancellationRequested { get { lock (sync) return cancelElapsed.HasValue; } }
        internal ExcelDiffProgressStage CurrentStage { get { lock (sync) return currentStage; } }
        internal string CurrentOperation { get { lock (sync) return currentOperation; } }

        internal void SetCurrentStage(ExcelDiffProgressStage stage, string operation)
        {
            lock (sync)
            {
                currentStage = stage;
                currentOperation = operation ?? string.Empty;
            }
        }

        internal void RecordCancelRequested(ExcelDiffProgressStage stage, string operation)
        {
            lock (sync)
            {
                if (cancelElapsed.HasValue) return;
                cancelElapsed = stopwatch.Elapsed;
                Add(ExcelDiffCancellationEventKind.CancelRequested, stage, operation);
            }
        }

        internal void Record(ExcelDiffCancellationEventKind kind, ExcelDiffProgressStage stage, string operation)
        {
            lock (sync) Add(kind, stage, operation);
        }

        internal void RecordOnce(ExcelDiffCancellationEventKind kind, ExcelDiffProgressStage stage, string operation)
        {
            lock (sync)
            {
                foreach (ExcelDiffCancellationEvent item in events)
                {
                    if (item.Kind == kind) return;
                }
                Add(kind, stage, operation);
            }
        }

        internal IList<ExcelDiffCancellationEvent> Snapshot()
        {
            lock (sync) return new List<ExcelDiffCancellationEvent>(events);
        }

        private void Add(ExcelDiffCancellationEventKind kind, ExcelDiffProgressStage stage, string operation)
        {
            DateTime utc = DateTime.UtcNow;
            TimeSpan elapsed = stopwatch.Elapsed;
            TimeSpan? sinceCancel = cancelElapsed.HasValue ? elapsed - cancelElapsed.Value : (TimeSpan?)null;
            events.Add(new ExcelDiffCancellationEvent(kind, taskId, stage, operation, utc, elapsed, sinceCancel, Thread.CurrentThread.ManagedThreadId));
        }
    }

    internal interface IExcelDiffProgress
    {
        Guid TaskId { get; }
        ExcelDiffCancellationTrace CancellationTrace { get; }
        void Report(ExcelDiffProgressInfo progress);
    }

    /// <summary>只包含托管值的线程安全进度快照；绝不持有 Excel COM 或 WinForms 控件。</summary>
    internal sealed class ExcelDiffProgressSnapshot
    {
        internal Guid TaskId { get; private set; }
        internal ExcelDiffProgressStage Stage { get; private set; }
        internal string Operation { get; private set; }
        internal string SheetName { get; private set; }
        internal int SheetIndex { get; private set; }
        internal int SheetCount { get; private set; }
        internal long Processed { get; private set; }
        internal long? Total { get; private set; }
        internal int Differences { get; private set; }
        internal int AddedRows { get; private set; }
        internal int DeletedRows { get; private set; }
        internal int ChangedRows { get; private set; }
        internal ExcelDiffProgressStatus Status { get; private set; }
        internal TimeSpan Elapsed { get; private set; }
        internal TimeSpan StageElapsed { get; private set; }
        internal DateTime LastProgressUtc { get; private set; }
        internal TimeSpan SinceLastProgress { get; private set; }

        internal static ExcelDiffProgressSnapshot Create(
            Guid taskId,
            ExcelDiffProgressInfo progress,
            TimeSpan elapsed,
            TimeSpan stageElapsed,
            DateTime lastProgressUtc,
            DateTime nowUtc)
        {
            return new ExcelDiffProgressSnapshot
            {
                TaskId = taskId,
                Stage = progress.Stage,
                Operation = progress.Operation,
                SheetName = progress.SheetName,
                SheetIndex = progress.CompletedSheets,
                SheetCount = progress.TotalSheets,
                Processed = progress.Processed,
                Total = progress.Total,
                Differences = progress.Differences,
                AddedRows = progress.AddedRows,
                DeletedRows = progress.DeletedRows,
                ChangedRows = progress.ChangedRows,
                Status = progress.Status,
                Elapsed = elapsed,
                StageElapsed = stageElapsed,
                LastProgressUtc = lastProgressUtc,
                SinceLastProgress = nowUtc - lastProgressUtc
            };
        }
    }

    /// <summary>
    /// 独立计时和心跳组件。Timer 回调只复制锁内托管状态，比较线程仍负责所有 Excel COM 操作。
    /// </summary>
    internal sealed class ExcelDiffProgressMonitor : IDisposable, IExcelDiffProgress
    {
        private readonly object sync = new object();
        private readonly Action<ExcelDiffProgressSnapshot> onSnapshot;
        private readonly Guid taskId = Guid.NewGuid();
        private readonly Stopwatch stopwatch = new Stopwatch();
        private readonly Stopwatch stageStopwatch = new Stopwatch();
        private Timer timer;
        private ExcelDiffProgressInfo latest;
        private DateTime lastProgressUtc;
        private bool disposed;
        private bool cancellationRequested;

        internal ExcelDiffProgressMonitor(Action<ExcelDiffProgressSnapshot> onSnapshot)
        {
            this.onSnapshot = onSnapshot;
            CancellationTrace = new ExcelDiffCancellationTrace(taskId);
            latest = new ExcelDiffProgressInfo(
                ExcelDiffProgressStage.Starting, "等待比较线程启动", string.Empty, 0, 0, 0, null, 0, 0, 0, 0,
                ExcelDiffProgressStatus.Running);
        }

        internal Guid TaskId { get { return taskId; } }
        internal ExcelDiffCancellationTrace CancellationTrace { get; private set; }

        Guid IExcelDiffProgress.TaskId { get { return taskId; } }
        ExcelDiffCancellationTrace IExcelDiffProgress.CancellationTrace { get { return CancellationTrace; } }

        internal void Start()
        {
            lock (sync)
            {
                if (disposed || stopwatch.IsRunning) return;
                stopwatch.Start();
                stageStopwatch.Start();
                lastProgressUtc = DateTime.UtcNow;
                timer = new Timer(Tick, null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
            }
        }

        public void Report(ExcelDiffProgressInfo progress)
        {
            if (progress == null) return;
            lock (sync)
            {
                if (disposed) return;
                if (!stopwatch.IsRunning) Start();
                if (cancellationRequested && progress.Status == ExcelDiffProgressStatus.Running) return;
                if (latest.Stage != progress.Stage)
                {
                    stageStopwatch.Restart();
                }
                latest = progress;
                lastProgressUtc = DateTime.UtcNow;
            }
        }

        internal void RequestCancellation()
        {
            ExcelDiffProgressStage stage;
            lock (sync)
            {
                if (disposed || cancellationRequested) return;
                cancellationRequested = true;
                stage = latest.Stage;
                latest = new ExcelDiffProgressInfo(
                    stage,
                    "已收到取消请求，正在停止比较",
                    latest.SheetName,
                    latest.CompletedSheets,
                    latest.TotalSheets,
                    latest.Processed,
                    latest.Total,
                    latest.Differences,
                    latest.AddedRows,
                    latest.DeletedRows,
                    latest.ChangedRows,
                    ExcelDiffProgressStatus.Cancelling);
                lastProgressUtc = DateTime.UtcNow;
            }
            CancellationTrace.RecordCancelRequested(stage, "已收到取消请求，正在停止比较");
            Tick(null);
        }

        internal void Stop(ExcelDiffProgressStatus status, ExcelDiffProgressStage stage, string operation)
        {
            Report(new ExcelDiffProgressInfo(stage, operation, latest.SheetName, latest.CompletedSheets, latest.TotalSheets,
                latest.Processed, latest.Total, latest.Differences, latest.AddedRows, latest.DeletedRows, latest.ChangedRows, status));
            lock (sync)
            {
                if (!disposed) stopwatch.Stop();
            }
            Tick(null);
        }

        private void Tick(object state)
        {
            ExcelDiffProgressSnapshot snapshot;
            Action<ExcelDiffProgressSnapshot> callback;
            lock (sync)
            {
                if (disposed || !stopwatch.IsRunning && latest.Status == ExcelDiffProgressStatus.Running) return;
                DateTime now = DateTime.UtcNow;
                snapshot = ExcelDiffProgressSnapshot.Create(taskId, latest, stopwatch.Elapsed, stageStopwatch.Elapsed, lastProgressUtc, now);
                callback = onSnapshot;
            }
            if (callback != null)
            {
                try { callback(snapshot); }
                catch { /* 监控回调不能影响比较线程或 Timer */ }
            }
        }

        public void Dispose()
        {
            Timer oldTimer;
            lock (sync)
            {
                if (disposed) return;
                disposed = true;
                stopwatch.Stop();
                stageStopwatch.Stop();
                oldTimer = timer;
                timer = null;
            }
            if (oldTimer != null) oldTimer.Dispose();
        }
    }
}
