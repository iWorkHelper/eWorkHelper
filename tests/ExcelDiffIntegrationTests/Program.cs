using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace eWorkhelper
{
    internal static class Program
    {
        private static Excel.Application application;
        private static Excel.Workbook baseline;
        private static Excel.Workbook candidate;
        private static string tempDirectory;
        private static readonly HashSet<int> excelProcessesBefore = new HashSet<int>();
        private static readonly IList<int> ownedExcelProcessIds = new List<int>();
        private static string lifecycleFailure;

        [STAThread]
        private static void Main()
        {
            foreach (int processId in GetExcelProcessIds()) excelProcessesBefore.Add(processId);
            Exception failure = null;
            try
            {
                RunIntegrationTest();
            }
            catch (Exception exception)
            {
                failure = exception;
            }
            finally
            {
                CloseOwnedWorkbook(ref candidate);
                CloseOwnedWorkbook(ref baseline);
                if (application != null)
                {
                    CloseAllOwnedWorkbooks(application);
                    application.Quit();
                    ComHelper.Release(application);
                    application = null;
                }
                // 仅用于集成测试诊断隐式 Interop RCW；正式业务代码不依赖 GC 回收 COM。
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                lifecycleFailure = WaitForOwnedExcelProcessesToExit();
                if (!string.IsNullOrEmpty(tempDirectory) && Directory.Exists(tempDirectory))
                {
                    Directory.Delete(tempDirectory, true);
                }
            }

            if (failure != null)
            {
                Console.Error.WriteLine("EXCEL_DIFF_INTEGRATION_FAIL type=" + failure.GetType().FullName);
                Console.Error.WriteLine("message=" + failure.Message);
                Console.Error.WriteLine(failure.StackTrace);
                Environment.ExitCode = 1;
                return;
            }

            if (lifecycleFailure != null)
            {
                Console.Error.WriteLine("EXCEL_DIFF_INTEGRATION_FAIL type=ExcelProcessLifecycle");
                Console.Error.WriteLine("message=" + lifecycleFailure);
                Environment.ExitCode = 1;
                return;
            }

            Console.WriteLine("EXCEL_DIFF_INTEGRATION_PASS differences=7 sheets=2 source_hashes_unchanged=true comment_append=true merge_preserved=true conflict_rejected=true cancellation_safe=true existing_features_regression=true owned_excel_exit=true");
        }

        private static void RunIntegrationTest()
        {
            tempDirectory = Path.Combine(Path.GetTempPath(), "eWorkHelper-ExcelDiff-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempDirectory);
            string baselinePath = Path.Combine(tempDirectory, "baseline.xlsx");
            string candidatePath = Path.Combine(tempDirectory, "candidate.xlsx");
            string outputPath = Path.Combine(tempDirectory, "result.xlsx");

            application = new Excel.Application();
            application.Visible = false;
            application.DisplayAlerts = false;
            CaptureOwnedExcelProcesses();
            baseline = CreateSampleWorkbook(baselinePath, false);
            baseline.Close(false);
            ComHelper.Release(baseline);
            baseline = null;
            candidate = CreateSampleWorkbook(candidatePath, true);
            ExcelDiffWorkbookIdentity smokeBaseline = ExcelDiffWorkbookIdentity.CaptureActive(application);
            AssertEqual(Path.GetFullPath(candidatePath), smokeBaseline.Path, "active workbook captured at entry");
            ExcelDiffForm guiSmoke = new ExcelDiffForm(application, smokeBaseline);
            guiSmoke.Close();
            guiSmoke.Dispose();
            candidate.Close(false);
            ComHelper.Release(candidate);
            candidate = null;

            string baselineHash = HashFile(baselinePath);
            string candidateHash = HashFile(candidatePath);
            ExcelDiffRequest request = CreateRequest(baselinePath, candidatePath, outputPath);
            ExcelDiffService service = new ExcelDiffService(application);
            ExcelDiffExecutionResult result = service.Compare(request, null, CancellationToken.None);
            AssertEqual(ExcelDiffOperationStatus.Completed, result.Status, "completed status");
            AssertEqual(7, result.DifferenceCount, "difference count");
            AssertEqual(0, result.AddedRowCount, "added row count");
            AssertEqual(0, result.DeletedRowCount, "deleted row count");
            // Data 的第 1～3 行及 Second 的第 1 行均包含差异。
            AssertEqual(4, result.ChangedRowCount, "changed row count across both sheets");
            ExcelDiffResult dataResult = result.SheetResults[0].Result;
            ExcelDiffCell dateDifference = FindDifference(dataResult, "R3C1");
            ExcelDiffCell errorDifference = FindDifference(dataResult, "R1C5");
            AssertEqual(ExcelCellSourceKind.Date, dateDifference.BaselineValue.SourceKind, "date source kind");
            AssertEqual(ExcelCellSourceKind.Formula, errorDifference.BaselineValue.SourceKind, "formula error source kind");
            AssertTrue(errorDifference.BaselineValue.IsError, "formula error metadata");
            AssertTrue(FindDifferenceOrNull(dataResult, "R2C5") == null, "equal formula results are equal");
            AssertTrue(File.Exists(outputPath), "output exists");

            VerifyOutput(outputPath);
            AssertEqual(baselineHash, HashFile(baselinePath), "baseline hash");
            AssertEqual(candidateHash, HashFile(candidatePath), "candidate hash");

            request.Options.OverwriteExistingOutput = false;
            ExcelDiffExecutionResult conflict = service.Compare(request, null, CancellationToken.None);
            AssertEqual(ExcelDiffOperationStatus.Failed, conflict.Status, "output conflict status");

            string cancelledPath = Path.Combine(tempDirectory, "cancelled.xlsx");
            request = CreateRequest(baselinePath, candidatePath, cancelledPath);
            CancellationTokenSource cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            ExcelDiffExecutionResult cancelled = service.Compare(request, null, cancellation.Token);
            AssertEqual(ExcelDiffOperationStatus.Cancelled, cancelled.Status, "cancel status");
            AssertTrue(!File.Exists(cancelledPath), "cancelled output absent");

            string midOperationCancelledPath = Path.Combine(tempDirectory, "mid-operation-cancelled.xlsx");
            request = CreateRequest(baselinePath, candidatePath, midOperationCancelledPath);
            CancellationTokenSource midOperationCancellation = new CancellationTokenSource();
            ExcelDiffExecutionResult midOperationCancelled = service.Compare(
                request,
                new CancelOnStageProgress(midOperationCancellation, ExcelDiffProgressStage.Comparing),
                midOperationCancellation.Token);
            AssertEqual(ExcelDiffOperationStatus.Cancelled, midOperationCancelled.Status, "mid-operation cancel status");
            AssertTrue(!File.Exists(midOperationCancelledPath), "mid-operation cancelled output absent");
            AssertCancellationTrace(midOperationCancelled.CancellationTrace);

            RunWorkbookIdentityRegression(baselinePath, candidatePath);
            RunExistingFeatureRegression();
        }

        private static void AssertCancellationTrace(ExcelDiffCancellationTrace trace)
        {
            IList<ExcelDiffCancellationEvent> events = trace == null ? null : trace.Snapshot();
            AssertTrue(events != null && events.Count >= 5, "cancellation trace has lifecycle events");
            ExcelDiffCancellationEventKind[] expected =
            {
                ExcelDiffCancellationEventKind.CancelRequested,
                ExcelDiffCancellationEventKind.CancelObserved,
                ExcelDiffCancellationEventKind.BusinessStopped,
                ExcelDiffCancellationEventKind.CleanupStarted,
                ExcelDiffCancellationEventKind.CleanupCompleted
            };
            int next = 0;
            foreach (ExcelDiffCancellationEvent item in events)
            {
                if (next < expected.Length && item.Kind == expected[next]) next++;
            }
            AssertEqual(expected.Length, next, "cancellation trace order");
        }

        private static void RunWorkbookIdentityRegression(string baselinePath, string candidatePath)
        {
            Excel.Workbooks workbooks = null;
            Excel.Workbook fixedBaseline = null;
            Excel.Workbook openCandidate = null;
            Excel.Workbook unsaved = null;
            try
            {
                AssertRejected(delegate { ExcelDiffWorkbookIdentity.CaptureActive(application); }, "no active workbook");
                workbooks = application.Workbooks;
                unsaved = workbooks.Add();
                unsaved.Activate();
                AssertRejected(delegate { ExcelDiffWorkbookIdentity.CaptureActive(application); }, "unsaved workbook");
                unsaved.Close(false);
                ComHelper.Release(unsaved);
                unsaved = null;

                fixedBaseline = workbooks.Open(baselinePath, UpdateLinks: 0, ReadOnly: false, AddToMru: false);
                fixedBaseline.Activate();
                ExcelDiffWorkbookIdentity identity = ExcelDiffWorkbookIdentity.CaptureActive(application);
                AssertEqual(Path.GetFullPath(baselinePath), identity.Path, "baseline fixed path");
                identity.Validate(application);
                AssertRejected(delegate { ExcelDiffWorkbookIdentity.ValidateCandidate(application, baselinePath.ToUpperInvariant(), identity.Path); }, "same actual file");
                string sameNameDirectory = Path.Combine(tempDirectory, "same-name");
                Directory.CreateDirectory(sameNameDirectory);
                string otherSameName = Path.Combine(sameNameDirectory, Path.GetFileName(baselinePath));
                File.Copy(baselinePath, otherSameName);
                AssertTrue(!ExcelDiffWorkbookIdentity.SamePath(otherSameName, baselinePath), "different directory is distinct path");
                AssertRejected(delegate { ExcelDiffWorkbookIdentity.ValidateCandidate(application, otherSameName, identity.Path); }, "Excel cannot open different files with same name");

                openCandidate = workbooks.Open(candidatePath, UpdateLinks: 0, ReadOnly: false, AddToMru: false);
                openCandidate.Activate();
                identity.Validate(application);
                ExcelDiffWorkbookIdentity.ValidateCandidate(application, candidatePath, identity.Path);
                VerifyFixedBaselineForm(identity, baselinePath, candidatePath);
                SetTestDirty(openCandidate);
                AssertRejected(delegate { ExcelDiffWorkbookIdentity.ValidateCandidate(application, candidatePath, identity.Path); }, "dirty candidate");
                openCandidate.Close(false);
                ComHelper.Release(openCandidate);
                openCandidate = null;

                SetTestDirty(fixedBaseline);
                AssertRejected(delegate { identity.Validate(application); }, "dirty baseline");
                fixedBaseline.Save();
                AssertRejected(delegate { identity.Validate(application); }, "baseline changed and saved after window opened");
                fixedBaseline.Close(false);
                ComHelper.Release(fixedBaseline);
                fixedBaseline = null;
                AssertRejected(delegate { identity.Validate(application); }, "closed baseline");
                fixedBaseline = workbooks.Open(baselinePath, UpdateLinks: 0, ReadOnly: false, AddToMru: false);
                AssertRejected(delegate { identity.Validate(application); }, "reopened baseline");
            }
            finally
            {
                if (unsaved != null) unsaved.Close(false);
                if (openCandidate != null) openCandidate.Close(false);
                if (fixedBaseline != null) fixedBaseline.Close(false);
                ComHelper.Release(unsaved);
                ComHelper.Release(openCandidate);
                ComHelper.Release(fixedBaseline);
                ComHelper.Release(workbooks);
            }
        }

        private static void AssertRejected(Action action, string message)
        {
            try { action(); }
            catch (InvalidOperationException) { return; }
            throw new InvalidOperationException("expected rejection: " + message);
        }

        private static void SetTestDirty(Excel.Workbook workbook)
        {
            Excel.Sheets sheets = null;
            Excel.Worksheet sheet = null;
            try
            {
                sheets = workbook.Worksheets;
                sheet = sheets[1] as Excel.Worksheet;
                SetCellValue(sheet, 8, 1, "not saved");
            }
            finally
            {
                ComHelper.Release(sheet);
                ComHelper.Release(sheets);
            }
        }

        private static void VerifyFixedBaselineForm(ExcelDiffWorkbookIdentity identity, string baselinePath, string candidatePath)
        {
            using (ExcelDiffForm form = new ExcelDiffForm(application, identity))
            {
                Label baselineDisplay = (Label)GetFormField(form, "baselineNameLabel");
                ComboBox candidateChoice = (ComboBox)GetFormField(form, "candidateComboBox");
                DataGridView mappings = (DataGridView)GetFormField(form, "mappingGrid");
                AssertEqual(Path.GetFileName(baselinePath), baselineDisplay.Text, "baseline shown by filename only");
                AssertTrue(form.Controls.Find("baselinePathTextBox", true).Length == 0 &&
                    typeof(ExcelDiffForm).GetField("baselinePathTextBox", BindingFlags.NonPublic | BindingFlags.Instance) == null,
                    "baseline path display removed");
                AssertEqual(Path.GetFullPath(candidatePath), candidateChoice.SelectedItem.ToString(), "candidate selected path");
                AssertTrue(fixedBaselineNotClosed(application, baselinePath), "sheet enumeration must not close baseline");
                VerifyFormLayoutAndStates(form, baselinePath);

                MethodInfo refresh = typeof(ExcelDiffForm).GetMethod("RefreshSheets", BindingFlags.NonPublic | BindingFlags.Instance);
                refresh.Invoke(form, null);
                ComboBox baselineSheets = (ComboBox)GetFormField(form, "baselineSheetComboBox");
                ComboBox candidateSheets = (ComboBox)GetFormField(form, "candidateSheetComboBox");
                AssertTrue(baselineSheets.Items.Contains("Data") && baselineSheets.Items.Contains("Second"), "baseline sheet list");
                AssertTrue(candidateSheets.Items.Contains("Data") && candidateSheets.Items.Contains("Second"), "candidate sheet list");
                mappings.Rows.Add(1, "Data", "Data");
                mappings.Rows.Add(2, "Second", "Second");
                AssertTrue(!((Label)GetFormField(form, "emptyMappingLabel")).Visible, "mapping hint hidden after first mapping");
                object[] args = { null };
                MethodInfo build = typeof(ExcelDiffForm).GetMethod("BuildRequest", BindingFlags.NonPublic | BindingFlags.Instance);
                ExcelDiffRequest request = (ExcelDiffRequest)build.Invoke(form, args);
                AssertTrue(request != null, "form builds request: " + args[0]);
                AssertEqual(Path.GetFullPath(baselinePath), request.Baseline.FilePath, "form baseline direction");
                AssertEqual(Path.GetFullPath(candidatePath), request.Candidate.FilePath, "form candidate direction");
                AssertEqual(2, request.SheetPairs.Count, "form mapping count");
                AssertTrue(request.Options.AddBaselineComment, "GUI defaults to baseline comment");

                string anotherPath = Path.Combine(tempDirectory, "another-candidate.xlsx");
                File.Copy(candidatePath, anotherPath);
                Type choiceType = typeof(ExcelDiffForm).GetNestedType("ExcelWorkbookChoice", BindingFlags.NonPublic);
                object anotherChoice = Activator.CreateInstance(
                    choiceType, BindingFlags.NonPublic | BindingFlags.Instance, null, new object[] { anotherPath }, null);
                candidateChoice.Items.Add(anotherChoice);
                candidateChoice.SelectedItem = anotherChoice;
                AssertEqual(0, mappings.Rows.Count, "candidate switch clears stale mappings");
                AssertTrue(((Label)GetFormField(form, "emptyMappingLabel")).Visible, "mapping hint restored after clearing");
                AssertTrue(baselineSheets.Items.Contains("Data"), "baseline retained after candidate switch");
                form.Hide();
            }
        }

        private static void VerifyFormLayoutAndStates(ExcelDiffForm form, string baselinePath)
        {
            form.ShowInTaskbar = false;
            form.Opacity = 0;
            form.Show();
            System.Windows.Forms.Application.DoEvents();
            using (Graphics graphics = form.CreateGraphics())
                Console.WriteLine("GUI_DPI_X=" + graphics.DpiX);
            {
                Label name = (Label)GetFormField(form, "baselineNameLabel");
                Label hint = (Label)GetFormField(form, "emptyMappingLabel");
                TextBox details = (TextBox)GetFormField(form, "statusDetails");
                CheckBox comment = (CheckBox)GetFormField(form, "commentCheckBox");
                CheckBox highlight = (CheckBox)GetFormField(form, "highlightCheckBox");
                Panel preview = (Panel)GetFormField(form, "colorPreview");
                Button color = (Button)GetFormField(form, "colorButton");
                Button start = (Button)GetFormField(form, "startButton");
                Button cancel = (Button)GetFormField(form, "cancelButton");
                Button open = (Button)GetFormField(form, "openResultButton");
                Button close = (Button)GetFormField(form, "closeButton");
                Label stage = (Label)GetFormField(form, "stageLabel");
                ProgressBar progress = (ProgressBar)GetFormField(form, "progressBar");
                DataGridView grid = (DataGridView)GetFormField(form, "mappingGrid");
                AssertEqual(Path.GetFileName(baselinePath), name.Text, "baseline filename emphasis");
                AssertTrue(hint.Visible && hint.Bounds.Top > grid.ColumnHeadersHeight, "empty mapping hint below headers");
                AssertTrue(details.Multiline && details.ScrollBars == ScrollBars.Vertical, "long status supports scrolling");
                AssertTrue(color.BackColor != Color.Red && preview.BackColor == Color.Red, "separate color preview and action");
                AssertTrue(highlight.Checked && comment.Checked, "option defaults");
                AssertEqual("基准内容写入批注", comment.Text, "comment option label");
                AssertTrue(!cancel.Enabled && !open.Enabled && start.Enabled && close.Enabled, "idle buttons");
                int[] optionColumns = ((TableLayoutPanel)comment.Parent).GetColumnWidths();
                AssertTrue(TextRenderer.MeasureText(comment.Text, comment.Font).Width <= optionColumns[0] + optionColumns[1], "comment text fits");
                AssertFormGeometry(form, "normal");

                SaveLayoutScreenshot(form, "normal");

                string longError = new string('测', 500) + Environment.NewLine + "末尾诊断";
                details.Text = longError;
                AssertTrue(details.Text.EndsWith("末尾诊断", StringComparison.Ordinal), "long error retained");
                form.Size = form.MinimumSize;
                System.Windows.Forms.Application.DoEvents();
                AssertFormGeometry(form, "minimum");

                // Scale 模拟控件布局压力，不等同于切换真实 Windows DPI。
                form.Scale(new SizeF(1.25f, 1.25f));
                System.Windows.Forms.Application.DoEvents();
                AssertFormGeometry(form, "simulated-125-percent");
                SaveLayoutScreenshot(form, "simulated-125");
                form.Scale(new SizeF(1.2f, 1.2f));
                System.Windows.Forms.Application.DoEvents();
                AssertFormGeometry(form, "simulated-150-percent");
                SaveLayoutScreenshot(form, "simulated-150");

                ComboBox baselineSheets = (ComboBox)GetFormField(form, "baselineSheetComboBox");
                ComboBox candidateSheets = (ComboBox)GetFormField(form, "candidateSheetComboBox");
                Button addMapping = (Button)GetFormField(form, "addMappingButton");
                Button removeMapping = (Button)GetFormField(form, "removeMappingButton");
                string longSheet = new string('长', 31);
                baselineSheets.Items.Add(longSheet);
                candidateSheets.Items.Add(longSheet);
                baselineSheets.SelectedItem = longSheet;
                candidateSheets.SelectedItem = longSheet;
                addMapping.PerformClick();
                AssertTrue(grid.Rows.Count == 1 && !hint.Visible, "single mapping hides empty state");
                AssertEqual(longSheet, grid.Rows[0].Cells[1].ToolTipText, "long sheet name tooltip");
                grid.Rows[0].Selected = true;
                removeMapping.PerformClick();
                AssertTrue(grid.Rows.Count == 0 && hint.Visible, "removing all mappings restores hint");
                for (int index = 0; index < 20; index++) grid.Rows.Add(index + 1, "基准" + index, "候选" + index);
                AssertTrue(grid.ScrollBars == ScrollBars.Both && grid.DisplayedRowCount(false) < 20, "multiple mappings scroll");
                grid.Rows.Clear();

                highlight.Checked = false;
                AssertTrue(!color.Enabled && preview.BackColor == Color.Red, "disabled highlight retains color");
                highlight.Checked = true;
                AssertTrue(color.Enabled, "color button restored");

                CancellationTokenSource cancellation = new CancellationTokenSource();
                typeof(ExcelDiffForm).GetField("cancellationSource", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, cancellation);
                typeof(ExcelDiffForm).GetField("busy", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, true);
                MethodInfo setBusy = typeof(ExcelDiffForm).GetMethod("SetBusyState", BindingFlags.NonPublic | BindingFlags.Instance);
                setBusy.Invoke(form, new object[] { true });
                AssertTrue(!start.Enabled && cancel.Enabled && !close.Enabled && !open.Enabled && progress.Visible, "running buttons and progress");
                cancel.PerformClick();
                AssertTrue(cancellation.IsCancellationRequested && !cancel.Enabled && stage.Text == "正在停止比较…", "cancel request state");
                typeof(ExcelDiffForm).GetField("busy", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, false);
                setBusy.Invoke(form, new object[] { false });
                AssertTrue(start.Enabled && !cancel.Enabled && close.Enabled && !progress.Visible, "buttons restored after cancellation");
                cancellation.Dispose();
                typeof(ExcelDiffForm).GetField("cancellationSource", BindingFlags.NonPublic | BindingFlags.Instance).SetValue(form, null);

                MethodInfo applyResult = typeof(ExcelDiffForm).GetMethod("ApplyExecutionResult", BindingFlags.NonPublic | BindingFlags.Instance);
                IList<ExcelDiffSheetExecutionResult> sheets = new List<ExcelDiffSheetExecutionResult>();
                string resultPath = Path.Combine(tempDirectory, "result.xlsx");
                applyResult.Invoke(form, new object[] { ExcelDiffExecutionResult.Completed(resultPath, new ExcelDiffLog(), sheets, 3) });
                setBusy.Invoke(form, new object[] { false });
                AssertTrue(stage.Text == "比较完成" && open.Enabled && details.Text.Contains(resultPath), "completed result and open entry");
                open.PerformClick();
                Excel.Workbook openedResult = ExcelDiffWorkbookIdentity.FindOpenWorkbook(application, resultPath);
                AssertTrue(openedResult != null, "result entry opens generated workbook");
                try { openedResult.Close(false); }
                finally { ComHelper.ReleaseBorrowed(openedResult); }
                applyResult.Invoke(form, new object[] { ExcelDiffExecutionResult.Completed(resultPath, new ExcelDiffLog(), sheets, 0) });
                AssertTrue(stage.Text == "未发现差异", "zero-difference state");
                applyResult.Invoke(form, new object[] { ExcelDiffExecutionResult.Cancelled(new ExcelDiffLog(), sheets, 0) });
                setBusy.Invoke(form, new object[] { false });
                AssertTrue(stage.Text == "已取消" && !open.Enabled, "cancelled result clears old output");
                applyResult.Invoke(form, new object[] { ExcelDiffExecutionResult.Failed(new ExcelDiffLog(), sheets, 0, new InvalidOperationException(longError)) });
                AssertTrue(stage.Text == "失败" && details.Text.EndsWith("末尾诊断", StringComparison.Ordinal), "failure details not truncated");
            }
        }

        private static void SaveLayoutScreenshot(ExcelDiffForm form, string state)
        {
            string destination = Environment.GetEnvironmentVariable("IWORKHELPER_GUI_SCREENSHOT");
            if (string.IsNullOrWhiteSpace(destination)) return;
            string path = state == "normal" ? destination : Path.Combine(
                Path.GetDirectoryName(destination), Path.GetFileNameWithoutExtension(destination) + "-" + state + ".png");
            using (Bitmap bitmap = new Bitmap(form.Width, form.Height))
            {
                form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            }
            Console.WriteLine("GUI_SCREENSHOT=" + path);
        }

        private static void AssertFormGeometry(ExcelDiffForm form, string state)
        {
            DataGridView grid = (DataGridView)GetFormField(form, "mappingGrid");
            TextBox details = (TextBox)GetFormField(form, "statusDetails");
            CheckBox comment = (CheckBox)GetFormField(form, "commentCheckBox");
            Button start = (Button)GetFormField(form, "startButton");
            Button close = (Button)GetFormField(form, "closeButton");
            Rectangle gridBounds = form.RectangleToClient(grid.RectangleToScreen(grid.ClientRectangle));
            Rectangle detailBounds = form.RectangleToClient(details.RectangleToScreen(details.ClientRectangle));
            Rectangle startBounds = form.RectangleToClient(start.RectangleToScreen(start.ClientRectangle));
            Rectangle closeBounds = form.RectangleToClient(close.RectangleToScreen(close.ClientRectangle));
            AssertTrue(gridBounds.Height >= 50 && gridBounds.Bottom < detailBounds.Top, state + " mapping and status geometry");
            AssertTrue(detailBounds.Height >= 35 && detailBounds.Bottom <= startBounds.Top, state + " details and actions geometry");
            AssertTrue(startBounds.Right <= form.ClientSize.Width && closeBounds.Right <= form.ClientSize.Width, state + " buttons inside client");
            AssertTrue(TextRenderer.MeasureText(comment.Text, comment.Font).Height <= comment.Height, state + " comment not clipped");
        }

        private static bool fixedBaselineNotClosed(Excel.Application excelApplication, string path)
        {
            Excel.Workbook workbook = ExcelDiffWorkbookIdentity.FindOpenWorkbook(excelApplication, path);
            try { return workbook != null; }
            finally { ComHelper.ReleaseBorrowed(workbook); }
        }

        private static object GetFormField(ExcelDiffForm form, string name)
        {
            return typeof(ExcelDiffForm).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form);
        }

        private static void CloseAllOwnedWorkbooks(Excel.Application excelApplication)
        {
            Excel.Workbooks workbooks = null;
            try
            {
                workbooks = excelApplication.Workbooks;
                Console.WriteLine("EXCEL_FINAL_WORKBOOKS=" + workbooks.Count);
                for (int index = workbooks.Count; index >= 1; index--)
                {
                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = workbooks[index];
                        workbook.Close(false);
                    }
                    finally
                    {
                        ComHelper.Release(workbook);
                    }
                }
            }
            finally
            {
                ComHelper.Release(workbooks);
            }
        }

        private static void CaptureOwnedExcelProcesses()
        {
            for (int attempt = 0; attempt < 30 && ownedExcelProcessIds.Count == 0; attempt++)
            {
                foreach (int processId in GetExcelProcessIds())
                {
                    if (!excelProcessesBefore.Contains(processId) && !ownedExcelProcessIds.Contains(processId))
                    {
                        ownedExcelProcessIds.Add(processId);
                    }
                }

                if (ownedExcelProcessIds.Count == 0) Thread.Sleep(100);
            }

            Console.WriteLine("EXCEL_OWNED_PIDS=" + (ownedExcelProcessIds.Count == 0
                ? "none"
                : string.Join(",", ownedExcelProcessIds)));
        }

        private static string WaitForOwnedExcelProcessesToExit()
        {
            if (ownedExcelProcessIds.Count == 0)
            {
                return null;
            }

            for (int attempt = 0; attempt < 50; attempt++)
            {
                bool anyRunning = false;
                foreach (int processId in ownedExcelProcessIds)
                {
                    if (Process.GetProcessesByName("EXCEL").Any(process => process.Id == processId))
                    {
                        anyRunning = true;
                        break;
                    }
                }

                if (!anyRunning)
                {
                    Console.WriteLine("EXCEL_OWNED_EXITED=" + string.Join(",", ownedExcelProcessIds));
                    return null;
                }

                Thread.Sleep(100);
            }

            return "测试创建的 Excel PID 未退出：" + string.Join(",", ownedExcelProcessIds);
        }

        private static IList<int> GetExcelProcessIds()
        {
            IList<int> processIds = new List<int>();
            foreach (Process process in Process.GetProcessesByName("EXCEL"))
            {
                try { processIds.Add(process.Id); }
                finally { process.Dispose(); }
            }
            return processIds;
        }

        private static void RunExistingFeatureRegression()
        {
            Excel.Workbook workbook = null;
            Excel.Sheets worksheets = null;
            Excel.Worksheet sheet = null;
            Excel.Range filterRange = null;
            Excel.Range activeCell = null;
            Excel.Range hiddenRow = null;
            Excel.ListObject table = null;
            Excel.Range merged = null;
            Excel.Range mergedSelection = null;
            try
            {
                workbook = application.Workbooks.Add();
                worksheets = workbook.Worksheets;
                sheet = worksheets[1] as Excel.Worksheet;
                sheet.Name = "回归";
                application.Visible = true;
                workbook.Activate();
                sheet.Activate();

                SetCellValue(sheet, 1, 1, "类型");
                SetCellValue(sheet, 1, 2, "值");
                SetCellValue(sheet, 2, 1, "保留");
                SetCellValue(sheet, 2, 2, "A");
                SetCellValue(sheet, 3, 1, "过滤");
                SetCellValue(sheet, 3, 2, "B");
                filterRange = sheet.Range["A1:B3"];
                table = sheet.ListObjects.Add(
                    Excel.XlListObjectSourceType.xlSrcRange,
                    filterRange,
                    Type.Missing,
                    Excel.XlYesNoGuess.xlYes,
                    Type.Missing);
                activeCell = sheet.Range["A2"];
                activeCell.Select();

                BatchFilterService filterService = new BatchFilterService();
                BatchFilterContext filterContext;
                string filterError;
                AssertTrue(filterService.TryCreateContext(application, out filterContext, out filterError),
                    "batch filter context: " + filterError);
                AssertTrue(filterContext != null, "batch filter context object");
                try
                {
                    BatchFilterResult filterResult = filterService.Apply(
                        filterContext,
                        new List<string> { "保留" },
                        BatchFilterMatchMode.Equals,
                        delegate { return false; },
                        null);
                    AssertTrue(filterResult.Applied, "batch filter applied");
                    hiddenRow = sheet.Rows[3] as Excel.Range;
                    AssertTrue(Convert.ToBoolean(hiddenRow.Hidden), "batch filter hides unmatched row");
                    ComHelper.Release(hiddenRow);
                    hiddenRow = null;
                    AssertTrue(filterService.Clear(filterContext), "batch filter clear");
                    hiddenRow = sheet.Rows[3] as Excel.Range;
                    AssertTrue(!Convert.ToBoolean(hiddenRow.Hidden), "batch filter clear restores row");
                }
                finally
                {
                    filterContext.Dispose();
                }

                merged = sheet.Range["D5:E5"];
                merged.Merge();
                SetRangeTopLeftValue(merged, "合并值");
                mergedSelection = sheet.Range["D5:E5"];
                mergedSelection.Select();
                int processed = new UnmergeAndFillService().Execute(application);
                AssertEqual(1, processed, "unmerge processed area count");
                AssertTrue(!Convert.ToBoolean(mergedSelection.MergeCells), "unmerge removes merge");
                AssertEqual("合并值", GetCellValue(sheet, 5, 4) as string, "unmerge top-left value");
                AssertEqual("合并值", GetCellValue(sheet, 5, 5) as string, "unmerge fills second cell");
            }
            finally
            {
                ComHelper.Release(hiddenRow);
                ComHelper.Release(table);
                ComHelper.Release(mergedSelection);
                ComHelper.Release(merged);
                ComHelper.Release(activeCell);
                ComHelper.Release(filterRange);
                ComHelper.Release(sheet);
                ComHelper.Release(worksheets);
                if (workbook != null) workbook.Close(false);
                ComHelper.Release(workbook);
            }
        }

        private sealed class CancelOnStageProgress : IExcelDiffProgress
        {
            private readonly CancellationTokenSource source;
            private readonly ExcelDiffProgressStage stage;
            private readonly Guid taskId = Guid.NewGuid();
            private readonly ExcelDiffCancellationTrace cancellationTrace;

            internal CancelOnStageProgress(CancellationTokenSource source, ExcelDiffProgressStage stage)
            {
                this.source = source;
                this.stage = stage;
                cancellationTrace = new ExcelDiffCancellationTrace(taskId);
            }

            public Guid TaskId { get { return taskId; } }
            public ExcelDiffCancellationTrace CancellationTrace { get { return cancellationTrace; } }

            public void Report(ExcelDiffProgressInfo progress)
            {
                if (progress.Stage == stage) source.Cancel();
            }
        }

        private static ExcelDiffRequest CreateRequest(string baselinePath, string candidatePath, string outputPath)
        {
            IList<ExcelDiffSheetPair> pairs = new List<ExcelDiffSheetPair>
            {
                new ExcelDiffSheetPair("Data", "Data", "Data pair"),
                new ExcelDiffSheetPair("Second", "Second", "Second pair")
            };
            ExcelDiffOptions options = new ExcelDiffOptions
            {
                HighlightDifferences = true,
                AddBaselineComment = true,
                HighlightColor = Color.Yellow,
                OverwriteExistingOutput = false,
                ValidateOutputAfterSave = true
            };
            return new ExcelDiffRequest(
                ExcelDiffWorkbookSource.FromFile(baselinePath),
                ExcelDiffWorkbookSource.FromFile(candidatePath),
                pairs,
                options,
                new ExcelDiffOutputOptions(outputPath));
        }

        private static Excel.Workbook CreateSampleWorkbook(string path, bool candidateFile)
        {
            Excel.Workbook workbook = null;
            Excel.Workbooks workbooks = null;
            Excel.Worksheet sheet = null;
            Excel.Worksheet second = null;
            Excel.Sheets sheets = null;
            Excel.Range merge = null;
            try
            {
                workbooks = application.Workbooks;
                workbook = workbooks.Add();
                sheets = workbook.Worksheets;
                sheet = sheets[1] as Excel.Worksheet;
                sheet.Name = "Data";
                second = sheets.Add(After: sheet) as Excel.Worksheet;
                second.Name = "Second";
                SetCellValue(sheet, 1, 1, candidateFile ? "new" : "old");
                SetCellValue(sheet, 1, 2, 10);
                SetCellNumberFormat(sheet, 1, 2, "0.00");
                SetCellValue(sheet, 2, 1, candidateFile ? null : "removed");
                SetCellValue(sheet, 2, 2, candidateFile ? "added" : null);
                SetCellValue(
                    sheet,
                    3,
                    1,
                    (candidateFile ? new DateTime(2026, 9, 19) : new DateTime(2026, 9, 18)).ToOADate());
                SetCellNumberFormat(sheet, 3, 1, "yyyy-mm-dd");
                SetCellFormula(sheet, 1, 5, candidateFile ? "=NA()" : "=1/0");
                SetCellFormula(sheet, 2, 5, candidateFile ? "=2" : "=1+1");
                merge = sheet.Range["C1:D1"];
                merge.Merge();
                SetRangeTopLeftValue(merge, candidateFile ? "new-merge" : "old-merge");
                SetColumnWidth(sheet, 2, 18);
                SetRowHeight(sheet, 1, 24);
                SetCellBold(sheet, 1, 2, true);
                SetCellColor(sheet, 1, 2, Color.LightBlue);
                if (candidateFile)
                {
                    AddCellComment(sheet, 1, 1, "existing comment");
                }
                SetCellValue(second, 1, 1, "same");
                SetCellValue(second, 1, 2, candidateFile ? "candidate" : "baseline");
                workbook.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook);
                return workbook;
            }
            finally
            {
                ComHelper.Release(merge);
                ComHelper.Release(second);
                ComHelper.Release(sheet);
                ComHelper.Release(sheets);
                ComHelper.Release(workbooks);
            }
        }

        private static void VerifyOutput(string path)
        {
            Excel.Workbook workbook = null;
            Excel.Workbooks workbooks = null;
            Excel.Worksheet data = null;
            Excel.Worksheet second = null;
            Excel.Sheets sheets = null;
            Excel.Comment comment = null;
            Excel.Range mergeRange = null;
            try
            {
                workbooks = application.Workbooks;
                workbook = workbooks.Open(path, UpdateLinks: 0, ReadOnly: true, AddToMru: false);
                sheets = workbook.Worksheets;
                data = sheets["Data"] as Excel.Worksheet;
                second = sheets["Second"] as Excel.Worksheet;
                AssertEqual("CHG", GetCellValue(data, 1, 1) as string, "changed row status");
                AssertEqual("new", GetCellValue(data, 1, 2) as string, "candidate value shifted");
                AssertEqual("added", GetCellValue(data, 2, 3) as string, "added value shifted");
                AssertTrue(GetCellColor(data, 1, 2) != 0d, "highlight shifted");
                mergeRange = data.Range["D1:E1"];
                AssertTrue(Convert.ToBoolean(mergeRange.MergeCells), "merge preserved");
                comment = GetCellComment(data, 1, 2);
                AssertTrue(comment != null, "comment exists");
                AssertTrue(comment.Text().IndexOf("existing comment", StringComparison.Ordinal) >= 0, "existing comment preserved");
                AssertTrue(comment.Text().IndexOf("[ExcelDiff]", StringComparison.Ordinal) >= 0, "diff comment appended");
                AssertEqual("candidate", GetCellValue(second, 1, 3) as string, "second sheet shifted");
            }
            finally
            {
                ComHelper.Release(comment);
                ComHelper.Release(mergeRange);
                if (workbook != null) workbook.Close(false);
                ComHelper.Release(second);
                ComHelper.Release(data);
                ComHelper.Release(sheets);
                ComHelper.Release(workbooks);
                ComHelper.Release(workbook);
            }
        }

        private static string HashFile(string path)
        {
            using (SHA256 sha256 = SHA256.Create())
            using (FileStream stream = File.OpenRead(path))
            {
                return BitConverter.ToString(sha256.ComputeHash(stream));
            }
        }

        private static void SetCellValue(Excel.Worksheet sheet, int row, int column, object value)
        {
            Excel.Range cell = null;
            try { cell = sheet.Cells[row, column]; cell.Value2 = value; }
            finally { ComHelper.Release(cell); }
        }

        private static void SetCellNumberFormat(Excel.Worksheet sheet, int row, int column, string format)
        {
            Excel.Range cell = null;
            try { cell = sheet.Cells[row, column]; cell.NumberFormat = format; }
            finally { ComHelper.Release(cell); }
        }

        private static void SetCellFormula(Excel.Worksheet sheet, int row, int column, string formula)
        {
            Excel.Range cell = null;
            try { cell = sheet.Cells[row, column]; cell.Formula = formula; }
            finally { ComHelper.Release(cell); }
        }

        private static void SetCellColor(Excel.Worksheet sheet, int row, int column, Color color)
        {
            Excel.Range cell = null;
            Excel.Interior interior = null;
            try { cell = sheet.Cells[row, column]; interior = cell.Interior; interior.Color = ColorTranslator.ToOle(color); }
            finally { ComHelper.Release(interior); ComHelper.Release(cell); }
        }

        private static void SetCellBold(Excel.Worksheet sheet, int row, int column, bool bold)
        {
            Excel.Range cell = null;
            Excel.Font font = null;
            try { cell = sheet.Cells[row, column]; font = cell.Font; font.Bold = bold; }
            finally { ComHelper.Release(font); ComHelper.Release(cell); }
        }

        private static void SetColumnWidth(Excel.Worksheet sheet, int column, double width)
        {
            Excel.Range range = null;
            try { range = sheet.Columns[column] as Excel.Range; range.ColumnWidth = width; }
            finally { ComHelper.Release(range); }
        }

        private static void SetRowHeight(Excel.Worksheet sheet, int row, double height)
        {
            Excel.Range range = null;
            try { range = sheet.Rows[row] as Excel.Range; range.RowHeight = height; }
            finally { ComHelper.Release(range); }
        }

        private static void SetRangeTopLeftValue(Excel.Range range, object value)
        {
            Excel.Range topLeft = null;
            try { topLeft = range.Cells[1, 1] as Excel.Range; topLeft.Value2 = value; }
            finally { ComHelper.Release(topLeft); }
        }

        private static void AddCellComment(Excel.Worksheet sheet, int row, int column, string text)
        {
            Excel.Range cell = null;
            Excel.Comment comment = null;
            try { cell = sheet.Cells[row, column]; comment = cell.AddComment(text); }
            finally { ComHelper.Release(comment); ComHelper.Release(cell); }
        }

        private static object GetCellValue(Excel.Worksheet sheet, int row, int column)
        {
            Excel.Range cell = null;
            try { cell = sheet.Cells[row, column]; return cell.Value2; }
            finally { ComHelper.Release(cell); }
        }

        private static double GetCellColor(Excel.Worksheet sheet, int row, int column)
        {
            Excel.Range cell = null;
            Excel.Interior interior = null;
            try { cell = sheet.Cells[row, column]; interior = cell.Interior; return Convert.ToDouble(interior.Color); }
            finally { ComHelper.Release(interior); ComHelper.Release(cell); }
        }

        private static Excel.Comment GetCellComment(Excel.Worksheet sheet, int row, int column)
        {
            Excel.Range cell = null;
            try { cell = sheet.Cells[row, column]; return cell.Comment; }
            finally { ComHelper.Release(cell); }
        }

        private static ExcelDiffCell FindDifference(ExcelDiffResult result, string coordinate)
        {
            ExcelDiffCell difference = FindDifferenceOrNull(result, coordinate);
            if (difference == null) throw new InvalidOperationException("missing difference " + coordinate);
            return difference;
        }

        private static ExcelDiffCell FindDifferenceOrNull(ExcelDiffResult result, string coordinate)
        {
            foreach (ExcelDiffCell difference in result.Differences)
            {
                if (difference.Coordinate.ToString() == coordinate) return difference;
            }
            return null;
        }

        private static void CloseOwnedWorkbook(ref Excel.Workbook workbook)
        {
            if (workbook == null) return;
            try { workbook.Close(false); } finally { ComHelper.Release(workbook); workbook = null; }
        }

        private static void AssertTrue(bool value, string message)
        {
            if (!value) throw new InvalidOperationException(message);
        }

        private static void AssertEqual<T>(T expected, T actual, string message)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual))
            {
                throw new InvalidOperationException(string.Format("{0}: expected {1}, actual {2}", message, expected, actual));
            }
        }
    }
}
