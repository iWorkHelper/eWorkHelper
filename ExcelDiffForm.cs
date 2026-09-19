using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace eWorkhelper
{
    /// <summary>宿主线程上捕获的基准身份；不持有宿主 Workbook RCW。</summary>
    internal sealed class ExcelDiffWorkbookIdentity
    {
        private readonly long comIdentity;
        private readonly string fileHash;

        private ExcelDiffWorkbookIdentity(string path, long comIdentity, string fileHash)
        {
            Path = path;
            this.comIdentity = comIdentity;
            this.fileHash = fileHash;
        }

        internal string Path { get; private set; }

        internal static ExcelDiffWorkbookIdentity CaptureActive(Excel.Application application)
        {
            if (application == null) throw new InvalidOperationException("无法获取当前 Excel 实例。");
            Excel.Workbook workbook = null;
            try
            {
                workbook = application.ActiveWorkbook;
                if (workbook == null) throw new InvalidOperationException("没有活动工作簿。请先打开要作为基准的工作簿。");
                string path = workbook.FullName;
                if (string.IsNullOrEmpty(workbook.Path) || !File.Exists(path))
                    throw new InvalidOperationException("当前工作簿尚未保存或文件不存在。请先保存后重试。");
                if (!workbook.Saved) throw new InvalidOperationException("当前工作簿有未保存修改。请先保存后重试。");
                EnsureSupported(path);
                return new ExcelDiffWorkbookIdentity(System.IO.Path.GetFullPath(path), GetComIdentity(workbook), HashFile(path));
            }
            finally { ComHelper.ReleaseBorrowed(workbook); }
        }

        internal void Validate(Excel.Application application)
        {
            Excel.Workbook workbook = FindOpenWorkbook(application, Path);
            if (workbook == null) throw new InvalidOperationException("基准工作簿已关闭或路径已改变。请从目标工作簿重新打开差异对比。");
            try
            {
                if (GetComIdentity(workbook) != comIdentity)
                    throw new InvalidOperationException("基准工作簿已被重新打开。请从目标工作簿重新打开差异对比。");
                if (!workbook.Saved)
                    throw new InvalidOperationException("基准工作簿有未保存修改，请先保存并重新打开差异对比。");
            }
            finally { ComHelper.ReleaseBorrowed(workbook); }
            if (!File.Exists(Path) || !string.Equals(HashFile(Path), fileHash, StringComparison.Ordinal))
                throw new InvalidOperationException("基准文件在窗口打开后发生变化，请从目标工作簿重新打开差异对比。");
        }

        internal static void ValidateCandidate(Excel.Application application, string candidatePath, string baselinePath)
        {
            if (string.IsNullOrWhiteSpace(candidatePath) || !File.Exists(candidatePath))
                throw new InvalidOperationException("待比较工作簿不存在，请重新选择。");
            EnsureSupported(candidatePath);
            if (SamePath(candidatePath, baselinePath))
                throw new InvalidOperationException("待比较工作簿不能与当前基准工作簿为同一文件。");
            if (string.Equals(System.IO.Path.GetFileName(candidatePath), System.IO.Path.GetFileName(baselinePath), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("两个工作簿虽然位于不同目录，但文件名相同；Excel 无法在同一实例同时打开。请先给待比较文件另存为不同名称后重试。");
            Excel.Workbook workbook = FindOpenWorkbook(application, candidatePath);
            try
            {
                if (workbook != null && !workbook.Saved)
                    throw new InvalidOperationException("待比较工作簿有未保存修改，请先保存后比较。");
            }
            finally { ComHelper.ReleaseBorrowed(workbook); }
        }

        internal static Excel.Workbook FindOpenWorkbook(Excel.Application application, string path)
        {
            Excel.Workbooks workbooks = null;
            try
            {
                workbooks = application.Workbooks;
                for (int index = 1; index <= workbooks.Count; index++)
                {
                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = workbooks[index];
                        if (SamePath(workbook.FullName, path))
                        {
                            Excel.Workbook found = workbook;
                            workbook = null;
                            return found;
                        }
                    }
                    finally { ComHelper.ReleaseBorrowed(workbook); }
                }
                return null;
            }
            finally { ComHelper.ReleaseBorrowed(workbooks); }
        }

        internal static bool SamePath(string left, string right)
        {
            return string.Equals(System.IO.Path.GetFullPath(left), System.IO.Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
        }

        private static void EnsureSupported(string path)
        {
            string extension = System.IO.Path.GetExtension(path);
            if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("仅支持 .xlsx 和 .xlsm 工作簿。");
        }

        private static long GetComIdentity(object workbook)
        {
            IntPtr pointer = Marshal.GetIUnknownForObject(workbook);
            try { return pointer.ToInt64(); }
            finally { Marshal.Release(pointer); }
        }

        private static string HashFile(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (SHA256 hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(stream));
        }
    }

    /// <summary>
    /// Excel 差异对比的 M3 WinForms 界面。宿主 COM 只在界面线程访问，
    /// 比较逻辑全部委托给 ExcelDiffService。
    /// </summary>
    internal sealed class ExcelDiffForm : Form
    {
        private readonly Excel.Application application;
        private readonly ExcelDiffWorkbookIdentity baseline;
        private readonly Label baselineNameLabel = new Label();
        private readonly ToolTip toolTip = new ToolTip();
        private readonly ComboBox candidateComboBox = new ComboBox();
        private readonly Button browseCandidateButton = new Button();
        private readonly ComboBox baselineSheetComboBox = new ComboBox();
        private readonly ComboBox candidateSheetComboBox = new ComboBox();
        private readonly DataGridView mappingGrid = new DataGridView();
        private readonly Label emptyMappingLabel = new Label();
        private readonly CheckBox highlightCheckBox = new CheckBox();
        private readonly Button colorButton = new Button();
        private readonly Panel colorPreview = new Panel();
        private readonly CheckBox commentCheckBox = new CheckBox();
        private readonly TextBox outputTextBox = new TextBox();
        private readonly Button browseOutputButton = new Button();
        private readonly Label stageLabel = new Label();
        private readonly Label countLabel = new Label();
        private readonly TextBox statusDetails = new TextBox();
        private readonly ProgressBar progressBar = new ProgressBar();
        private readonly Button addMappingButton = new Button();
        private readonly Button removeMappingButton = new Button();
        private readonly Button startButton = new Button();
        private readonly Button cancelButton = new Button();
        private readonly Button openResultButton = new Button();
        private readonly Button closeButton = new Button();
        private CancellationTokenSource cancellationSource;
        private ExcelDiffProgressMonitor progressMonitor;
        private Guid progressTaskId;
        private bool busy;
        private string lastResultPath;
        private Color highlightColor = Color.Red;

        internal ExcelDiffForm(Excel.Application application, ExcelDiffWorkbookIdentity baseline)
        {
            this.application = application ?? throw new ArgumentNullException("application");
            this.baseline = baseline ?? throw new ArgumentNullException("baseline");
            Text = "Excel 差异对比";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(780, 680);
            ClientSize = new Size(820, 710);
            FormBorderStyle = FormBorderStyle.Sizable;
            AutoScaleMode = AutoScaleMode.Font;
            KeyPreview = true;

            BuildLayout();
            baselineNameLabel.Text = Path.GetFileName(baseline.Path);
            Disposed += delegate
            {
                if (progressMonitor != null) progressMonitor.Dispose();
                if (cancellationSource != null) cancellationSource.Cancel();
                toolTip.Dispose();
            };
            LoadOpenChoices();
            candidateComboBox.SelectedIndexChanged += WorkbookSelectionChanged;
            UpdateCandidateTooltip();
            addMappingButton.Click += AddMappingButton_Click;
            removeMappingButton.Click += RemoveMappingButton_Click;
            browseCandidateButton.Click += delegate { BrowseWorkbook(); };
            browseOutputButton.Click += BrowseOutputButton_Click;
            colorButton.Click += ColorButton_Click;
            highlightCheckBox.CheckedChanged += delegate { colorButton.Enabled = !busy && highlightCheckBox.Checked; };
            startButton.Click += StartButton_Click;
            cancelButton.Click += CancelButton_Click;
            openResultButton.Click += delegate { OpenGeneratedResult(); };
            closeButton.Click += delegate { Close(); };
            FormClosing += ExcelDiffForm_FormClosing;
            Shown += delegate
            {
                RefreshSheets();
                baselineSheetComboBox.Focus();
            };
        }

        private void BuildLayout()
        {
            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(10),
                AutoScroll = true
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 124));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 142));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            Controls.Add(root);

            GroupBox files = new GroupBox { Text = "工作簿", Dock = DockStyle.Fill };
            TableLayoutPanel fileTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 3, Padding = new Padding(8) };
            fileTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            fileTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            fileTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 78));
            fileTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 23));
            fileTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 23));
            fileTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 27));
            baselineNameLabel.AutoEllipsis = true;
            baselineNameLabel.Dock = DockStyle.Fill;
            baselineNameLabel.TextAlign = ContentAlignment.MiddleLeft;
            baselineNameLabel.Font = new Font(Font, FontStyle.Bold);
            fileTable.Controls.Add(new Label { Text = "当前工作簿", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            fileTable.Controls.Add(baselineNameLabel, 1, 0);
            fileTable.SetColumnSpan(baselineNameLabel, 2);
            Label baselineHint = new Label { Text = "当前工作簿已固定为基准；更换请关闭窗口后重新启动。", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = SystemColors.GrayText };
            fileTable.Controls.Add(baselineHint, 1, 1);
            fileTable.SetColumnSpan(baselineHint, 2);
            AddFileRow(fileTable, 2, "待比较工作簿", candidateComboBox, browseCandidateButton, "浏览…");
            toolTip.SetToolTip(candidateComboBox, "选择已打开工作簿，或点击“浏览…”选择文件。");
            files.Controls.Add(fileTable);
            root.Controls.Add(files, 0, 0);

            GroupBox mappings = new GroupBox { Text = "Sheet 映射（基准 → 待比较，一对一）", Dock = DockStyle.Fill };
            TableLayoutPanel mappingTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 2, Padding = new Padding(8) };
            mappingTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            mappingTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            mappingTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82));
            mappingTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82));
            baselineSheetComboBox.Dock = DockStyle.Fill;
            candidateSheetComboBox.Dock = DockStyle.Fill;
            addMappingButton.Text = "添加映射";
            removeMappingButton.Text = "删除映射";
            addMappingButton.Dock = DockStyle.Fill;
            removeMappingButton.Dock = DockStyle.Fill;
            mappingTable.Controls.Add(new Label { Text = "基准 Sheet", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            mappingTable.Controls.Add(new Label { Text = "待比较 Sheet", AutoSize = true, Anchor = AnchorStyles.Left }, 1, 0);
            mappingTable.Controls.Add(baselineSheetComboBox, 0, 1);
            mappingTable.Controls.Add(candidateSheetComboBox, 1, 1);
            mappingTable.Controls.Add(addMappingButton, 2, 1);
            mappingTable.Controls.Add(removeMappingButton, 3, 1);
            mappings.Controls.Add(mappingTable);
            mappingTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));
            mappingTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            mappingGrid.Dock = DockStyle.Fill;
            mappingGrid.AllowUserToAddRows = false;
            mappingGrid.AllowUserToDeleteRows = false;
            mappingGrid.AllowUserToResizeRows = false;
            mappingGrid.AutoGenerateColumns = false;
            mappingGrid.MultiSelect = false;
            mappingGrid.ReadOnly = true;
            mappingGrid.RowHeadersVisible = false;
            mappingGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            mappingGrid.BackgroundColor = SystemColors.Window;
            mappingGrid.ShowCellToolTips = true;
            mappingGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "序号", Width = 55, AutoSizeMode = DataGridViewAutoSizeColumnMode.None, DataPropertyName = "Index" });
            mappingGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "基准 Sheet", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, DataPropertyName = "Baseline" });
            mappingGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "待比较 Sheet", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, DataPropertyName = "Candidate" });
            emptyMappingLabel.Text = "尚未添加映射。请选择两侧 Sheet，然后点击“添加映射”。";
            emptyMappingLabel.AutoSize = true;
            emptyMappingLabel.ForeColor = SystemColors.GrayText;
            emptyMappingLabel.BackColor = SystemColors.Window;
            emptyMappingLabel.Location = new Point(10, 38);
            mappingGrid.Controls.Add(emptyMappingLabel);
            mappingGrid.RowsAdded += delegate { UpdateMappingEmptyState(); };
            mappingGrid.RowsRemoved += delegate { UpdateMappingEmptyState(); };

            TableLayoutPanel mappingArea = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            mappingArea.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            mappingArea.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mappingArea.Controls.Add(mappingTable, 0, 0);
            mappingArea.Controls.Add(mappingGrid, 0, 1);
            mappings.Controls.Clear();
            mappings.Controls.Add(mappingArea);
            root.Controls.Add(mappings, 0, 1);
            UpdateMappingEmptyState();

            GroupBox options = new GroupBox { Text = "比较选项", Dock = DockStyle.Fill };
            TableLayoutPanel optionTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, RowCount = 3, Padding = new Padding(8) };
            optionTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 172));
            optionTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 34));
            optionTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 116));
            optionTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            optionTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 78));
            optionTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            optionTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            optionTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 29));
            highlightCheckBox.Text = "标记差异高亮";
            highlightCheckBox.Checked = true;
            highlightCheckBox.AutoSize = true;
            highlightCheckBox.Anchor = AnchorStyles.Left;
            colorPreview.BackColor = highlightColor;
            colorPreview.BorderStyle = BorderStyle.FixedSingle;
            colorPreview.Size = new Size(25, 20);
            colorPreview.Anchor = AnchorStyles.Left;
            toolTip.SetToolTip(colorPreview, "当前高亮颜色");
            colorButton.Text = "更改颜色…";
            colorButton.Dock = DockStyle.Fill;
            commentCheckBox.Text = "基准内容写入批注";
            commentCheckBox.Checked = true;
            commentCheckBox.AutoSize = true;
            commentCheckBox.Anchor = AnchorStyles.Left;
            outputTextBox.Dock = DockStyle.Fill;
            browseOutputButton.Text = "浏览…";
            browseOutputButton.Dock = DockStyle.Fill;
            optionTable.Controls.Add(highlightCheckBox, 0, 0);
            optionTable.Controls.Add(colorPreview, 1, 0);
            optionTable.Controls.Add(colorButton, 2, 0);
            optionTable.Controls.Add(commentCheckBox, 0, 1);
            optionTable.SetColumnSpan(commentCheckBox, 2);
            Label commentHint = new Label { Text = "在结果差异处记录基准值。", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = SystemColors.GrayText };
            optionTable.Controls.Add(commentHint, 2, 1);
            optionTable.SetColumnSpan(commentHint, 2);
            optionTable.Controls.Add(new Label { Text = "输出文件", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
            optionTable.Controls.Add(outputTextBox, 1, 2);
            optionTable.SetColumnSpan(outputTextBox, 3);
            optionTable.Controls.Add(browseOutputButton, 4, 2);
            options.Controls.Add(optionTable);
            root.Controls.Add(options, 0, 2);

            GroupBox status = new GroupBox { Text = "执行状态", Dock = DockStyle.Fill };
            TableLayoutPanel statusTable = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 3, Padding = new Padding(8) };
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 72));
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            statusTable.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            statusTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            statusTable.RowStyles.Add(new RowStyle(SizeType.Absolute, 24));
            statusTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            stageLabel.Text = "等待开始";
            stageLabel.Dock = DockStyle.Fill;
            stageLabel.TextAlign = ContentAlignment.MiddleLeft;
            countLabel.Text = "差异数量：0";
            countLabel.Dock = DockStyle.Fill;
            countLabel.TextAlign = ContentAlignment.MiddleLeft;
            statusDetails.Text = "选择待比较工作簿并添加 Sheet 映射后，即可开始比较。";
            statusDetails.ReadOnly = true;
            statusDetails.Multiline = true;
            statusDetails.WordWrap = true;
            statusDetails.ScrollBars = ScrollBars.Vertical;
            statusDetails.Dock = DockStyle.Fill;
            statusDetails.TabStop = false;
            progressBar.Dock = DockStyle.Fill;
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;
            progressBar.Visible = false;
            statusTable.Controls.Add(new Label { Text = "阶段", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            statusTable.Controls.Add(stageLabel, 1, 0);
            statusTable.Controls.Add(progressBar, 2, 0);
            statusTable.Controls.Add(new Label { Text = "统计", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            statusTable.Controls.Add(countLabel, 1, 1);
            statusTable.SetColumnSpan(countLabel, 2);
            statusTable.Controls.Add(new Label { Text = "详情", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
            statusTable.Controls.Add(statusDetails, 1, 2);
            statusTable.SetColumnSpan(statusDetails, 2);
            status.Controls.Add(statusTable);
            root.Controls.Add(status, 0, 3);

            FlowLayoutPanel buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
            startButton.Text = "开始比较";
            cancelButton.Text = "取消比较";
            openResultButton.Text = "打开结果";
            closeButton.Text = "关闭";
            startButton.Width = 116;
            startButton.Font = new Font(Font, FontStyle.Bold);
            cancelButton.Width = 92;
            openResultButton.Width = 92;
            closeButton.Width = 80;
            closeButton.Margin = new Padding(12, 3, 3, 3);
            startButton.TabIndex = 0;
            openResultButton.TabIndex = 1;
            cancelButton.TabIndex = 2;
            closeButton.TabIndex = 3;
            cancelButton.Enabled = false;
            openResultButton.Enabled = false;
            buttons.Controls.Add(closeButton);
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(openResultButton);
            buttons.Controls.Add(startButton);
            root.Controls.Add(buttons, 0, 4);
        }

        private static void AddFileRow(TableLayoutPanel table, int row, string label, ComboBox comboBox, Button browseButton, string buttonText)
        {
            comboBox.Dock = DockStyle.Fill;
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            browseButton.Text = buttonText;
            browseButton.Dock = DockStyle.Fill;
            table.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row);
            table.Controls.Add(comboBox, 1, row);
            table.Controls.Add(browseButton, 2, row);
        }

        private void LoadOpenChoices()
        {
            candidateComboBox.Items.Clear();
            Excel.Workbooks workbooks = null;
            try
            {
                workbooks = application.Workbooks;
                for (int index = 1; index <= workbooks.Count; index++)
                {
                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = workbooks[index];
                        string path = workbook.Path;
                        if (string.IsNullOrEmpty(path))
                        {
                            continue;
                        }

                        string displayPath = path + Path.DirectorySeparatorChar + workbook.Name;
                        if (!ExcelDiffWorkbookIdentity.SamePath(displayPath, baseline.Path))
                            candidateComboBox.Items.Add(new ExcelWorkbookChoice(displayPath));
                    }
                    catch
                    {
                        // 路径读取失败时忽略该项；临时取得的 Workbook 在 finally 中释放。
                    }
                    finally
                    {
                        ComHelper.ReleaseBorrowed(workbook);
                        workbook = null;
                    }
                }
            }
            finally
            {
                ComHelper.ReleaseBorrowed(workbooks);
            }

            if (candidateComboBox.Items.Count > 0) candidateComboBox.SelectedIndex = 0;
        }

        private void BrowseWorkbook()
        {
            if (busy) return;
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel 工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|所有文件 (*.*)|*.*";
                dialog.CheckFileExists = true;
                dialog.Multiselect = false;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                ExcelWorkbookChoice choice = new ExcelWorkbookChoice(dialog.FileName);
                candidateComboBox.Items.Add(choice);
                candidateComboBox.SelectedItem = choice;
            }
        }

        private void WorkbookSelectionChanged(object sender, EventArgs e)
        {
            if (busy) return;
            UpdateCandidateTooltip();
            ClearMappings();
            RefreshSheets();
        }

        private void UpdateCandidateTooltip()
        {
            ExcelWorkbookChoice choice = candidateComboBox.SelectedItem as ExcelWorkbookChoice;
            toolTip.SetToolTip(candidateComboBox, choice == null
                ? "选择已打开工作簿，或点击“浏览…”选择文件。"
                : choice.Path);
        }

        private void RefreshSheets()
        {
            baselineSheetComboBox.Items.Clear();
            candidateSheetComboBox.Items.Clear();
            ExcelWorkbookChoice candidate = candidateComboBox.SelectedItem as ExcelWorkbookChoice;
            try
            {
                baseline.Validate(application);
                foreach (string name in GetSheetNames(baseline.Path)) baselineSheetComboBox.Items.Add(name);
                if (baselineSheetComboBox.Items.Count > 0) baselineSheetComboBox.SelectedIndex = 0;
                if (candidate == null) return;
                ExcelDiffWorkbookIdentity.ValidateCandidate(application, candidate.Path, baseline.Path);
                foreach (string name in GetSheetNames(candidate.Path)) candidateSheetComboBox.Items.Add(name);
                if (candidateSheetComboBox.Items.Count > 0) candidateSheetComboBox.SelectedIndex = 0;
                SetDefaultOutput(candidate);
            }
            catch (Exception exception)
            {
                ShowError("读取 Sheet 列表失败：" + exception.Message);
            }
        }

        private IList<string> GetSheetNames(string path)
        {
            Excel.Workbook workbook = ExcelDiffWorkbookIdentity.FindOpenWorkbook(application, path);
            bool ownsWorkbook = false;
            Excel.Workbooks workbooks = null;
            Excel.Sheets sheets = null;
            try
            {
                if (workbook == null)
                {
                    workbooks = application.Workbooks;
                    workbook = workbooks.Open(path, UpdateLinks: 0, ReadOnly: true, AddToMru: false, IgnoreReadOnlyRecommended: true);
                    ownsWorkbook = true;
                }

                IList<string> names = new List<string>();
                sheets = workbook.Worksheets;
                for (int index = 1; index <= sheets.Count; index++)
                {
                    Excel.Worksheet sheet = null;
                    try
                    {
                        sheet = sheets[index] as Excel.Worksheet;
                        names.Add(sheet.Name);
                    }
                    finally
                    {
                        ComHelper.ReleaseBorrowed(sheet);
                    }
                }
                return names;
            }
            finally
            {
                ComHelper.Release(sheets);
                if (ownsWorkbook && workbook != null)
                {
                    try { workbook.Close(false); } finally { ComHelper.Release(workbook); }
                }
                else ComHelper.ReleaseBorrowed(workbook);
                ComHelper.ReleaseBorrowed(workbooks);
            }
        }

        private void SetDefaultOutput(ExcelWorkbookChoice candidate)
        {
            if (candidate == null || string.IsNullOrEmpty(candidate.Path)) return;
            string extension = Path.GetExtension(candidate.Path);
            string directory = Path.GetDirectoryName(candidate.Path);
            string name = Path.GetFileNameWithoutExtension(candidate.Path) + "_差异结果" + extension;
            outputTextBox.Text = Path.Combine(directory, name);
        }

        private void AddMappingButton_Click(object sender, EventArgs e)
        {
            string baseline = baselineSheetComboBox.SelectedItem as string;
            string candidate = candidateSheetComboBox.SelectedItem as string;
            if (string.IsNullOrEmpty(baseline) || string.IsNullOrEmpty(candidate))
            {
                ShowWarning("请先选择基准 Sheet 和待比较 Sheet。");
                return;
            }

            foreach (DataGridViewRow row in mappingGrid.Rows)
            {
                if (string.Equals(Convert.ToString(row.Cells[1].Value), baseline, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(Convert.ToString(row.Cells[2].Value), candidate, StringComparison.OrdinalIgnoreCase))
                {
                    ShowWarning("每个 Sheet 只能参与一组映射，请修改现有映射。");
                    return;
                }
            }
            int rowIndex = mappingGrid.Rows.Add(mappingGrid.Rows.Count + 1, baseline, candidate);
            mappingGrid.Rows[rowIndex].Cells[1].ToolTipText = baseline;
            mappingGrid.Rows[rowIndex].Cells[2].ToolTipText = candidate;
            UpdateMappingEmptyState();
        }

        private void RemoveMappingButton_Click(object sender, EventArgs e)
        {
            if (mappingGrid.SelectedRows.Count == 0) return;
            mappingGrid.Rows.RemoveAt(mappingGrid.SelectedRows[0].Index);
            for (int index = 0; index < mappingGrid.Rows.Count; index++) mappingGrid.Rows[index].Cells[0].Value = index + 1;
            UpdateMappingEmptyState();
        }

        private void ClearMappings()
        {
            mappingGrid.Rows.Clear();
            UpdateMappingEmptyState();
        }

        private void UpdateMappingEmptyState()
        {
            if (mappingGrid.IsDisposed) return;
            emptyMappingLabel.Visible = mappingGrid.Rows.Count == 0;
            removeMappingButton.Enabled = !busy && mappingGrid.Rows.Count > 0;
        }

        private void ColorButton_Click(object sender, EventArgs e)
        {
            if (busy) return;
            using (ColorDialog dialog = new ColorDialog { Color = highlightColor, FullOpen = true })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    highlightColor = dialog.Color;
                    colorPreview.BackColor = highlightColor;
                }
            }
        }

        private void BrowseOutputButton_Click(object sender, EventArgs e)
        {
            if (busy) return;
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "Excel 工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm";
                dialog.FileName = Path.GetFileName(outputTextBox.Text);
                dialog.OverwritePrompt = false;
                if (dialog.ShowDialog(this) == DialogResult.OK) outputTextBox.Text = dialog.FileName;
            }
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            if (busy) return;
            string validationError;
            ExcelDiffRequest request = BuildRequest(out validationError);
            if (request == null)
            {
                ShowWarning(validationError);
                return;
            }

            busy = true;
            cancellationSource = new CancellationTokenSource();
            progressMonitor = new ExcelDiffProgressMonitor(UpdateProgressSnapshot);
            progressTaskId = progressMonitor.TaskId;
            progressMonitor.Start();
            lastResultPath = null;
            openResultButton.Enabled = false;
            stageLabel.Text = "准备中";
            countLabel.Text = "差异数量：0";
            statusDetails.Text = "正在准备比较…";
            SetBusyState(true);
            try { StartDedicatedStaComparison(request, cancellationSource.Token); }
            catch (Exception exception)
            {
                cancellationSource.Dispose();
                cancellationSource = null;
                if (progressMonitor != null) progressMonitor.Dispose();
                progressMonitor = null;
                busy = false;
                SetBusyState(false);
                stageLabel.Text = "失败";
                ShowError("无法启动比较任务：" + exception.Message);
            }
        }

        private void StartDedicatedStaComparison(ExcelDiffRequest request, CancellationToken token)
        {
            Thread worker = new Thread(new ThreadStart(delegate
            {
                Excel.Application workerApplication = null;
                ExcelDiffExecutionResult result = null;
                Exception error = null;
                try
                {
                    // 只在线程内部创建并使用 Excel；不把宿主 Application 或 Workbook RCW 跨线程传入。
                    workerApplication = new Excel.Application();
                    workerApplication.Visible = false;
                    workerApplication.DisplayAlerts = false;
                    result = new ExcelDiffService(workerApplication).Compare(request, progressMonitor, token);
                    if (progressMonitor != null)
                        progressMonitor.Stop(result.Status == ExcelDiffOperationStatus.Completed
                            ? ExcelDiffProgressStatus.Completed
                            : result.Status == ExcelDiffOperationStatus.Cancelled ? ExcelDiffProgressStatus.Cancelled : ExcelDiffProgressStatus.Failed,
                            result.Status == ExcelDiffOperationStatus.Completed ? ExcelDiffProgressStage.Completed
                                : result.Status == ExcelDiffOperationStatus.Cancelled ? ExcelDiffProgressStage.Cancelled : ExcelDiffProgressStage.Failed,
                            result.Status == ExcelDiffOperationStatus.Completed ? "比较完成" : result.Status == ExcelDiffOperationStatus.Cancelled ? "比较已取消" : "比较失败");
                }
                catch (Exception exception)
                {
                    error = exception;
                    if (progressMonitor != null)
                        progressMonitor.Stop(ExcelDiffProgressStatus.Failed, ExcelDiffProgressStage.Failed, "比较失败");
                }
                finally
                {
                    if (workerApplication != null)
                    {
                        try { workerApplication.Quit(); }
                        catch
                        {
                            // 线程内 Excel 已由本次操作创建；退出异常不能让 UI 永远停留在忙碌状态。
                        }
                        finally { ComHelper.Release(workerApplication); }
                    }
                    if (progressMonitor != null)
                    {
                        progressMonitor.CancellationTrace.RecordOnce(
                            ExcelDiffCancellationEventKind.WorkerExited,
                            ExcelDiffProgressStage.CleaningUp,
                            "专用 STA 工作线程已退出");
                    }
                }

                CompleteDedicatedStaComparison(result, error);
            }));
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }

        private void CompleteDedicatedStaComparison(ExcelDiffExecutionResult result, Exception error)
        {
            if (IsDisposed || !IsHandleCreated) return;
            try
            {
                BeginInvoke((MethodInvoker)delegate
                {
                    try
                    {
                        if (error != null)
                        {
                            stageLabel.Text = "失败";
                            ShowError("执行比较失败：" + error.Message);
                        }
                        else if (result != null)
                        {
                            ShowExecutionResult(result);
                        }
                    }
                    finally
                    {
                        if (cancellationSource != null) cancellationSource.Dispose();
                        cancellationSource = null;
                        if (progressMonitor != null) progressMonitor.Dispose();
                        progressMonitor = null;
                        busy = false;
                        SetBusyState(false);
                    }
                });
            }
            catch (InvalidOperationException)
            {
                // 窗体正在关闭，关闭路径会阻止释放仍在执行的请求。
            }
        }

        private ExcelDiffRequest BuildRequest(out string error)
        {
            error = null;
            ExcelWorkbookChoice candidate = candidateComboBox.SelectedItem as ExcelWorkbookChoice;
            if (candidate == null) { error = "请选择待比较工作簿。"; return null; }
            if (mappingGrid.Rows.Count == 0) { error = "请至少添加一组 Sheet 映射。"; return null; }
            if (string.IsNullOrWhiteSpace(outputTextBox.Text)) { error = "请选择输出文件。"; return null; }
            try
            {
                baseline.Validate(application);
                ExcelDiffWorkbookIdentity.ValidateCandidate(application, candidate.Path, baseline.Path);
            }
            catch (Exception exception)
            {
                error = exception.Message;
                return null;
            }

            IList<ExcelDiffSheetPair> pairs = new List<ExcelDiffSheetPair>();
            foreach (DataGridViewRow row in mappingGrid.Rows)
            {
                pairs.Add(new ExcelDiffSheetPair(Convert.ToString(row.Cells[1].Value), Convert.ToString(row.Cells[2].Value), null));
            }
            ExcelDiffOptions options = new ExcelDiffOptions
            {
                HighlightDifferences = highlightCheckBox.Checked,
                HighlightColor = highlightColor,
                AddBaselineComment = commentCheckBox.Checked,
                OverwriteExistingOutput = false,
                ValidateOutputAfterSave = true
            };
            // 宿主工作簿只用于主线程状态校验；专用 STA 只接收文件路径。
            return new ExcelDiffRequest(
                ExcelDiffWorkbookSource.FromFile(baseline.Path),
                ExcelDiffWorkbookSource.FromFile(candidate.Path),
                pairs,
                options,
                new ExcelDiffOutputOptions(outputTextBox.Text.Trim()));
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            if (busy && cancellationSource != null && !cancellationSource.IsCancellationRequested)
            {
                if (progressMonitor != null) progressMonitor.RequestCancellation();
                cancellationSource.Cancel();
                stageLabel.Text = "正在停止比较…";
                statusDetails.Text = "已收到取消请求，正在停止后续比较；如当前处于 Excel 调用，将等待其安全返回。";
                cancelButton.Enabled = false;
            }
        }

        private void ShowExecutionResult(ExcelDiffExecutionResult result)
        {
            string message = ApplyExecutionResult(result);
            if (result.Status == ExcelDiffOperationStatus.Completed)
            {
                if (!string.IsNullOrEmpty(result.OutputPath))
                {
                    try { OpenResultWorkbook(result.OutputPath); }
                    catch (Exception exception) { message += Environment.NewLine + "结果已生成，但自动打开失败：" + exception.Message; }
                }
                statusDetails.Text = message;
                MessageBox.Show(this, message, "Excel 差异对比", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (result.Status == ExcelDiffOperationStatus.Cancelled)
            {
                MessageBox.Show(this, message, "Excel 差异对比", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                ShowError(message);
            }
        }

        private string ApplyExecutionResult(ExcelDiffExecutionResult result)
        {
            if (result.Status == ExcelDiffOperationStatus.Completed)
            {
                stageLabel.Text = result.DifferenceCount == 0 ? "未发现差异" : "比较完成";
                countLabel.Text = string.Format("差异单元格：{0}；新增行：{1}；删除行：{2}；修改行：{3}", result.DifferenceCount, result.AddedRowCount, result.DeletedRowCount, result.ChangedRowCount);
                lastResultPath = result.OutputPath;
                string message = result.DifferenceCount == 0 ? "未发现差异。" : "比较完成。";
                if (!string.IsNullOrEmpty(result.OutputPath)) message += Environment.NewLine + "结果文件：" + result.OutputPath;
                statusDetails.Text = message;
                return message;
            }
            lastResultPath = null;
            if (result.Status == ExcelDiffOperationStatus.Cancelled)
            {
                stageLabel.Text = "已取消";
                countLabel.Text = "差异数量：—（未生成结果）";
                string diagnostics = FormatCancellationDiagnostics(result.CancellationTrace);
                statusDetails.Text = "已取消比较；未生成结果文件，源工作簿未修改。" + diagnostics;
                return "比较已取消，源工作簿未修改。" + diagnostics;
            }
            stageLabel.Text = "失败";
            statusDetails.Text = FormatFailure(result);
            return statusDetails.Text;
        }

        private static string FormatCancellationDiagnostics(ExcelDiffCancellationTrace trace)
        {
            if (trace == null) return string.Empty;
            IList<ExcelDiffCancellationEvent> events = trace.Snapshot();
            ExcelDiffCancellationEvent requested = null;
            ExcelDiffCancellationEvent stopped = null;
            ExcelDiffCancellationEvent cleaned = null;
            ExcelDiffCancellationEvent exited = null;
            foreach (ExcelDiffCancellationEvent item in events)
            {
                if (item.Kind == ExcelDiffCancellationEventKind.CancelRequested) requested = item;
                else if (item.Kind == ExcelDiffCancellationEventKind.BusinessStopped) stopped = item;
                else if (item.Kind == ExcelDiffCancellationEventKind.CleanupCompleted) cleaned = item;
                else if (item.Kind == ExcelDiffCancellationEventKind.WorkerExited) exited = item;
            }
            if (requested == null) return string.Empty;
            string observed = stopped == null || !stopped.SinceCancel.HasValue ? "未知" : FormatDuration(stopped.SinceCancel.Value);
            string cleanup = cleaned == null || !cleaned.SinceCancel.HasValue ? "未知" : FormatDuration(cleaned.SinceCancel.Value);
            string worker = exited == null || !exited.SinceCancel.HasValue ? "未知" : FormatDuration(exited.SinceCancel.Value);
            return string.Format("{0}取消感知后业务停止：{1}；清理完成：{2}；工作线程退出：{3}。", Environment.NewLine, observed, cleanup, worker);
        }

        private void OpenGeneratedResult()
        {
            if (busy || string.IsNullOrEmpty(lastResultPath)) return;
            try { OpenResultWorkbook(lastResultPath); }
            catch (Exception exception) { ShowError("打开结果失败：" + exception.Message); }
        }

        private static string FormatFailure(ExcelDiffExecutionResult result)
        {
            if (result == null || result.Error == null) return "比较失败，请检查输入文件、Sheet 映射和输出路径。";
            return "比较失败：" + result.Error.Message;
        }

        private void OpenResultWorkbook(string path)
        {
            string fullPath = Path.GetFullPath(path);
            Excel.Workbooks workbooks = null;
            try
            {
                workbooks = application.Workbooks;
                for (int index = 1; index <= workbooks.Count; index++)
                {
                    Excel.Workbook workbook = null;
                    try
                    {
                        workbook = workbooks[index];
                        if (string.Equals(Path.GetFullPath(workbook.FullName), fullPath, StringComparison.OrdinalIgnoreCase))
                        {
                            workbook.Activate();
                            return;
                        }
                    }
                    finally
                    {
                        ComHelper.ReleaseBorrowed(workbook);
                    }
                }
                Excel.Workbook opened = workbooks.Open(fullPath, UpdateLinks: 0, ReadOnly: false, AddToMru: false, IgnoreReadOnlyRecommended: true);
                ComHelper.ReleaseBorrowed(opened);
            }
            finally
            {
                ComHelper.ReleaseBorrowed(workbooks);
            }
        }

        private void SetBusyState(bool isBusy)
        {
            startButton.Enabled = !isBusy;
            cancelButton.Enabled = isBusy && cancellationSource != null && !cancellationSource.IsCancellationRequested;
            openResultButton.Enabled = !isBusy && !string.IsNullOrEmpty(lastResultPath);
            closeButton.Enabled = !isBusy;
            candidateComboBox.Enabled = !isBusy;
            browseCandidateButton.Enabled = !isBusy;
            baselineSheetComboBox.Enabled = !isBusy;
            candidateSheetComboBox.Enabled = !isBusy;
            addMappingButton.Enabled = !isBusy;
            removeMappingButton.Enabled = !isBusy && mappingGrid.Rows.Count > 0;
            highlightCheckBox.Enabled = !isBusy;
            commentCheckBox.Enabled = !isBusy;
            colorButton.Enabled = !isBusy && highlightCheckBox.Checked;
            outputTextBox.ReadOnly = isBusy;
            browseOutputButton.Enabled = !isBusy;
            progressBar.Visible = isBusy;
        }

        private void ExcelDiffForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (busy)
            {
                e.Cancel = true;
                ShowWarning("比较正在执行，请先等待任务完成或请求取消。");
                return;
            }
        }

        private void UpdateProgressSnapshot(ExcelDiffProgressSnapshot snapshot)
        {
            if (snapshot == null || IsDisposed || !IsHandleCreated || snapshot.TaskId != progressTaskId) return;
            try
            {
                BeginInvoke((MethodInvoker)delegate
                {
                    if (IsDisposed || !busy || snapshot.TaskId != progressTaskId) return;
                    if (cancellationSource != null
                        && cancellationSource.IsCancellationRequested
                        && snapshot.Status == ExcelDiffProgressStatus.Running) return;
                    string sheet = string.IsNullOrEmpty(snapshot.SheetName) ? string.Empty : "；当前 Sheet：" + snapshot.SheetName;
                    string work = snapshot.Total.HasValue
                        ? string.Format("已完成 {0} / {1}", snapshot.Processed, snapshot.Total.Value)
                        : string.Format("已处理 {0}", snapshot.Processed);
                    stageLabel.Text = GetStageText(snapshot.Stage) + sheet;
                    countLabel.Text = string.Format("差异单元格：{0}；新增行：{1}；删除行：{2}；修改行：{3}", snapshot.Differences, snapshot.AddedRows, snapshot.DeletedRows, snapshot.ChangedRows);
                    string heartbeat = snapshot.SinceLastProgress >= TimeSpan.FromSeconds(10)
                        ? "；当前操作耗时较长，暂未收到新的进度"
                        : string.Empty;
                    statusDetails.Text = string.Format("{0}；{1}；Sheet：{2}/{3}；累计耗时：{4}；阶段耗时：{5}；最近进度：{6} 前{7}",
                        snapshot.Operation, work, snapshot.SheetIndex, snapshot.SheetCount,
                        FormatDuration(snapshot.Elapsed), FormatDuration(snapshot.StageElapsed),
                        FormatDuration(snapshot.SinceLastProgress), heartbeat);
                });
            }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        private static string FormatDuration(TimeSpan duration)
        {
            return string.Format("{0:00}:{1:00}:{2:00}", (int)duration.TotalHours, duration.Minutes, duration.Seconds);
        }

        private static string GetStageText(ExcelDiffProgressStage stage)
        {
            switch (stage)
            {
                case ExcelDiffProgressStage.CheckingParameters: return "正在检查工作簿和比较参数";
                case ExcelDiffProgressStage.InitializingExcel: return "正在初始化 Excel 比较环境";
                case ExcelDiffProgressStage.OpeningBaseline: return "正在打开基准工作簿";
                case ExcelDiffProgressStage.OpeningCandidate: return "正在打开待比较工作簿";
                case ExcelDiffProgressStage.LoadingSheets: return "正在加载 Sheet 信息";
                case ExcelDiffProgressStage.ReadingBaselineSheet: return "正在读取基准 Sheet";
                case ExcelDiffProgressStage.ReadingCandidateSheet: return "正在读取待比较 Sheet";
                case ExcelDiffProgressStage.ComparingCells: return "正在比较单元格";
                case ExcelDiffProgressStage.Comparing: return "正在比较单元格";
                case ExcelDiffProgressStage.Opening: return "正在打开工作簿";
                case ExcelDiffProgressStage.Reading: return "正在读取 Sheet";
                case ExcelDiffProgressStage.CopyingResult: return "正在创建结果副本";
                case ExcelDiffProgressStage.InsertingStatusColumn: return "正在插入行状态列";
                case ExcelDiffProgressStage.WritingRowStatuses: return "正在写入行状态";
                case ExcelDiffProgressStage.Marking: return "正在标记差异单元格";
                case ExcelDiffProgressStage.Saving: return "正在保存结果工作簿";
                case ExcelDiffProgressStage.CleaningUp: return "正在完成资源清理";
                case ExcelDiffProgressStage.Completed: return "比较完成";
                case ExcelDiffProgressStage.Cancelled: return "已取消";
                case ExcelDiffProgressStage.Failed: return "失败";
                default: return "准备中";
            }
        }

        private void ShowWarning(string message)
        {
            statusDetails.Text = message;
            MessageBox.Show(this, message, "Excel 差异对比", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShowError(string message)
        {
            statusDetails.Text = message;
            MessageBox.Show(this, message, "Excel 差异对比", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private sealed class ExcelWorkbookChoice
        {
            internal ExcelWorkbookChoice(string displayPath)
            {
                Path = displayPath;
            }

            internal string Path { get; private set; }

            public override string ToString()
            {
                return Path;
            }

        }

        private sealed class ExcelDiffProgress : IExcelDiffProgress
        {
            private readonly ExcelDiffForm form;
            private readonly Guid taskId = Guid.NewGuid();
            private readonly ExcelDiffCancellationTrace cancellationTrace;

            internal ExcelDiffProgress(ExcelDiffForm form)
            {
                this.form = form;
                cancellationTrace = new ExcelDiffCancellationTrace(taskId);
            }

            public Guid TaskId { get { return cancellationTrace.TaskId; } }
            public ExcelDiffCancellationTrace CancellationTrace { get { return cancellationTrace; } }

            public void Report(ExcelDiffProgressInfo progress)
            {
                if (form.IsDisposed || !form.IsHandleCreated) return;
                if (form.InvokeRequired)
                {
                    try
                    {
                        form.BeginInvoke((MethodInvoker)delegate { Report(progress); });
                    }
                    catch (Exception exception) when (exception is InvalidOperationException || exception is ObjectDisposedException)
                    {
                        // 窗体正在关闭，忽略已经无法显示的进度通知。
                    }
                    return;
                }
                if (!form.busy) return;
                bool cancellationPending = form.cancellationSource != null && form.cancellationSource.IsCancellationRequested;
                if (!cancellationPending || progress.Stage == ExcelDiffProgressStage.Cancelled)
                    form.stageLabel.Text = GetStageText(progress.Stage) + "（" + progress.CompletedSheets + "/" + progress.TotalSheets + "）";
                form.countLabel.Text = "差异数量：" + progress.Differences;
                if (!cancellationPending)
                    form.statusDetails.Text = "已处理 Sheet：" + progress.CompletedSheets + "/" + progress.TotalSheets;
                form.stageLabel.Refresh();
                form.countLabel.Refresh();
            }

            private static string GetStageText(ExcelDiffProgressStage stage)
            {
                switch (stage)
                {
                    case ExcelDiffProgressStage.Opening: return "打开工作簿";
                    case ExcelDiffProgressStage.Reading: return "读取 Sheet";
                    case ExcelDiffProgressStage.Comparing: return "比较内容";
                    case ExcelDiffProgressStage.Marking: return "标记差异";
                    case ExcelDiffProgressStage.Saving: return "保存结果";
                    case ExcelDiffProgressStage.Completed: return "已完成";
                    case ExcelDiffProgressStage.Cancelled: return "已取消";
                    case ExcelDiffProgressStage.Failed: return "失败";
                    default: return "准备中";
                }
            }
        }
    }
}
