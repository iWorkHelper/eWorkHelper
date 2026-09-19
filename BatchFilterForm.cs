using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;

namespace eWorkhelper
{
    internal sealed class BatchFilterForm : Form
    {
        private readonly BatchFilterService service;
        private readonly BatchFilterContext context;
        private readonly TextBox conditionsTextBox;
        private readonly RadioButton equalsRadioButton;
        private readonly RadioButton notEqualsRadioButton;
        private readonly RadioButton containsRadioButton;
        private readonly RadioButton notContainsRadioButton;
        private readonly Label resultLabel;
        private readonly Button applyButton;
        private readonly Button clearButton;

        // E-05：长任务期间的协作取消与忙碌标志。
        private readonly CancellationTokenSource cancellationSource = new CancellationTokenSource();
        private bool busy;

        internal BatchFilterForm(BatchFilterService service, BatchFilterContext context)
        {
            this.service = service;
            this.context = context;

            Text = "批量过滤";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            KeyPreview = true;
            ClientSize = new Size(520, 410);

            Label targetColumnLabel = new Label
            {
                AutoSize = true,
                Location = new Point(18, 18),
                Text = "当前过滤列：" + context.ColumnDisplayName
            };

            Label conditionsLabel = new Label
            {
                AutoSize = true,
                Location = new Point(18, 50),
                Text = "过滤条件（每行一个）："
            };

            conditionsTextBox = new TextBox
            {
                AcceptsReturn = true,
                AcceptsTab = false,
                Location = new Point(18, 73),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Size = new Size(484, 210)
            };

            Label matchModeLabel = new Label
            {
                AutoSize = true,
                Location = new Point(18, 292),
                Text = "匹配方式"
            };

            equalsRadioButton = new RadioButton
            {
                AutoSize = true,
                Location = new Point(18, 314),
                Text = "等于"
            };

            notEqualsRadioButton = new RadioButton
            {
                AutoSize = true,
                Location = new Point(92, 314),
                Text = "不等于"
            };

            containsRadioButton = new RadioButton
            {
                AutoSize = true,
                Checked = true,
                Location = new Point(180, 314),
                Text = "包含"
            };

            notContainsRadioButton = new RadioButton
            {
                AutoSize = true,
                Location = new Point(254, 314),
                Text = "不包含"
            };

            applyButton = new Button
            {
                Location = new Point(18, 347),
                Size = new Size(96, 30),
                Text = "应用过滤"
            };
            applyButton.Click += ApplyButton_Click;

            clearButton = new Button
            {
                Location = new Point(124, 347),
                Size = new Size(112, 30),
                Text = "清除本次过滤"
            };
            clearButton.Click += ClearButton_Click;

            Button closeButton = new Button
            {
                DialogResult = DialogResult.Cancel,
                Location = new Point(406, 347),
                Size = new Size(96, 30),
                Text = "关闭"
            };

            resultLabel = new Label
            {
                AutoEllipsis = true,
                Location = new Point(18, 383),
                Size = new Size(484, 20),
                Text = "请输入过滤条件。"
            };

            Controls.Add(targetColumnLabel);
            Controls.Add(conditionsLabel);
            Controls.Add(conditionsTextBox);
            Controls.Add(matchModeLabel);
            Controls.Add(equalsRadioButton);
            Controls.Add(notEqualsRadioButton);
            Controls.Add(containsRadioButton);
            Controls.Add(notContainsRadioButton);
            Controls.Add(applyButton);
            Controls.Add(clearButton);
            Controls.Add(closeButton);
            Controls.Add(resultLabel);
            CancelButton = closeButton;

            // E-04：构造函数中的 Excel 交互必须受保护；失败时降级为可手工输入的空白状态。
            try
            {
                BatchFilterInitialState initialState = service.LoadInitialState(context);
                conditionsTextBox.Text = string.Join(Environment.NewLine, initialState.Conditions);
                SetMatchMode(initialState.MatchMode);
                resultLabel.Text = initialState.StatusMessage;
            }
            catch (Exception ex)
            {
                conditionsTextBox.Text = string.Empty;
                SetMatchMode(BatchFilterMatchMode.Contains);
                resultLabel.Text = "无法读取当前筛选状态，请手工输入过滤条件。";
                MessageBox.Show(
                    "读取当前筛选状态失败：" + ex.Message,
                    "批量过滤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            Shown += delegate { conditionsTextBox.Focus(); };
            // E-05：Esc 在操作进行中触发取消，其余情况交由 CancelButton 处理。
            KeyDown += BatchFilterForm_KeyDown;
        }

        private void BatchFilterForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Escape || !busy)
            {
                return;
            }

            e.Handled = true;
            e.SuppressKeyPress = true;
            CancelOperation();
        }

        /// <summary>
        /// E-05：请求取消当前长任务。已取消后再次点击无副作用。
        /// </summary>
        private void CancelOperation()
        {
            if (busy && !cancellationSource.IsCancellationRequested)
            {
                cancellationSource.Cancel();
                resultLabel.Text = "正在取消…";
            }
        }

        private void SetMatchMode(BatchFilterMatchMode mode)
        {
            equalsRadioButton.Checked = mode == BatchFilterMatchMode.Equals;
            notEqualsRadioButton.Checked = mode == BatchFilterMatchMode.NotEquals;
            containsRadioButton.Checked = mode == BatchFilterMatchMode.Contains;
            notContainsRadioButton.Checked = mode == BatchFilterMatchMode.NotContains;
        }

        private void ApplyButton_Click(object sender, EventArgs e)
        {
            if (busy)
            {
                return;
            }

            IList<string> conditions = service.NormalizeConditions(conditionsTextBox.Text);
            if (conditions.Count == 0)
            {
                MessageBox.Show(
                    "请输入至少一个过滤条件。",
                    "批量过滤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            string validationError = BatchFilterService.ValidateConditions(conditions);
            if (validationError != null)
            {
                MessageBox.Show(
                    validationError,
                    "批量过滤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            BatchFilterMatchMode mode = equalsRadioButton.Checked
                ? BatchFilterMatchMode.Equals
                : notEqualsRadioButton.Checked
                    ? BatchFilterMatchMode.NotEquals
                    : containsRadioButton.Checked
                        ? BatchFilterMatchMode.Contains
                        : BatchFilterMatchMode.NotContains;

            RunLongOperation(
                delegate(BatchFilterCancellationCheck isCancelled, BatchFilterProgressReport report)
                {
                    return service.Apply(context, conditions, mode, isCancelled, report);
                },
                "应用过滤失败：");
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            if (busy)
            {
                return;
            }

            RunLongOperation(
                delegate(BatchFilterCancellationCheck isCancelled, BatchFilterProgressReport report)
                {
                    bool cleared = service.Clear(context);
                    return new BatchFilterResult
                    {
                        HasCounts = false,
                        Applied = true,
                        Message = cleared
                            ? "已清除本次过滤。"
                            : "当前目标字段没有由批量过滤应用的筛选。"
                    };
                },
                "清除过滤失败：");
        }

        /// <summary>
        /// E-05：统一的长任务外壳。Excel COM 调用始终留在 UI 线程上，
        /// 通过 Application.StatusBar/DoEvents 泵消息获得可取消的进度反馈。
        /// </summary>
        private void RunLongOperation(
            Func<BatchFilterCancellationCheck, BatchFilterProgressReport, BatchFilterResult> operation,
            string failurePrefix)
        {
            busy = true;
            applyButton.Enabled = false;
            clearButton.Enabled = false;
            Cursor previousCursor = Cursor;
            Cursor = Cursors.WaitCursor;

            try
            {
                BatchFilterProgressReport report = delegate(string message)
                {
                    if (IsDisposed)
                    {
                        return;
                    }

                    resultLabel.Text = message;
                    Application.DoEvents();
                };

                BatchFilterResult result = operation(
                    delegate { return cancellationSource.IsCancellationRequested; },
                    report);

                if (IsDisposed)
                {
                    return;
                }

                if (result.HasCounts)
                {
                    // E-06：明确区分“条件匹配行数”（目标列自身）与“当前实际可见行数”
                    //（还与其它字段的筛选取交集）。
                    resultLabel.Text = result.VisibleCount >= 0 && result.VisibleCount != result.MatchedRowCount
                        ? string.Format(
                            CultureInfo.CurrentCulture,
                            "共检查 {0} 行；本次条件匹配 {1} 行；当前实际显示 {2} 行（其它字段筛选会取交集）。",
                            result.CheckedCount,
                            result.MatchedRowCount,
                            result.VisibleCount)
                        : string.Format(
                            CultureInfo.CurrentCulture,
                            "共检查 {0} 行，匹配 {1} 行。",
                            result.CheckedCount,
                            result.MatchedRowCount);
                }
                else
                {
                    resultLabel.Text = result.Message;
                }
            }
            catch (OperationCanceledException)
            {
                if (!IsDisposed)
                {
                    resultLabel.Text = "操作已取消，工作表筛选状态已恢复。";
                }
            }
            catch (Exception ex)
            {
                // DoEvents 期间窗体若被释放，这里不再触碰已释放的控件。
                if (!IsDisposed)
                {
                    MessageBox.Show(
                        failurePrefix + ex.Message,
                        "批量过滤",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }
            finally
            {
                busy = false;
                if (!IsDisposed)
                {
                    Cursor = previousCursor;
                    applyButton.Enabled = true;
                    clearButton.Enabled = true;
                }
            }
        }

        /// <summary>
        /// E-03：释放 <see cref="BatchFilterContext"/> 持有的全部 Excel RCW。
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (busy)
                {
                    // 操作进行中：先请求取消，再释放上下文。
                    cancellationSource.Cancel();
                }

                BatchFilterContext ownedContext = context;
                if (ownedContext != null)
                {
                    ownedContext.Dispose();
                }
            }

            base.Dispose(disposing);
        }

        /// <summary>
        /// E-05：操作进行中不允许直接关闭窗体，避免 DoEvents 泵出的关闭消息导致已释放对象访问。
        /// </summary>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (busy)
            {
                e.Cancel = true;
                CancelOperation();
                return;
            }

            base.OnFormClosing(e);
        }
    }
}
