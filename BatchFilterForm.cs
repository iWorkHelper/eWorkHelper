using System;
using System.Collections.Generic;
using System.Drawing;
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

        internal BatchFilterForm(BatchFilterService service, BatchFilterContext context)
        {
            this.service = service;
            this.context = context;

            Text = "批量过滤";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
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

            Button applyButton = new Button
            {
                Location = new Point(18, 347),
                Size = new Size(96, 30),
                Text = "应用过滤"
            };
            applyButton.Click += ApplyButton_Click;

            Button clearButton = new Button
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

            BatchFilterInitialState initialState = service.LoadInitialState(context);
            conditionsTextBox.Text = string.Join(Environment.NewLine, initialState.Conditions);
            SetMatchMode(initialState.MatchMode);
            resultLabel.Text = initialState.StatusMessage;
            Shown += delegate { conditionsTextBox.Focus(); };
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

            try
            {
                BatchFilterMatchMode mode = equalsRadioButton.Checked
                    ? BatchFilterMatchMode.Equals
                    : notEqualsRadioButton.Checked
                        ? BatchFilterMatchMode.NotEquals
                        : containsRadioButton.Checked
                            ? BatchFilterMatchMode.Contains
                            : BatchFilterMatchMode.NotContains;
                BatchFilterResult result = service.Apply(context, conditions, mode);
                resultLabel.Text = string.Format(
                    "共检查 {0} 行，匹配 {1} 行。",
                    result.CheckedCount,
                    result.MatchedCount);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "应用过滤失败：" + ex.Message,
                    "批量过滤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            try
            {
                resultLabel.Text = service.Clear(context)
                    ? "已清除本次过滤。"
                    : "当前目标字段没有由批量过滤应用的筛选。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "清除过滤失败：" + ex.Message,
                    "批量过滤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            //
            // BatchFilterForm
            //
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Name = "BatchFilterForm";
            this.ResumeLayout(false);

        }
    }
}
