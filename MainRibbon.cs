using Microsoft.Office.Tools.Ribbon;
using System;
using System.Reflection;
using System.Windows.Forms;

namespace eWorkhelper
{
    public partial class MainRibbon
    {
        private readonly BatchFilterService batchFilterService = new BatchFilterService();
        private readonly UnmergeAndFillService unmergeAndFillService = new UnmergeAndFillService();

        private void btnBatchFilter_Click(object sender, RibbonControlEventArgs e)
        {
            BatchFilterContext context;
            string errorMessage;
            if (!batchFilterService.TryCreateContext(Globals.ThisAddIn.Application, out context, out errorMessage))
            {
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    MessageBox.Show(errorMessage, "批量过滤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                return;
            }

            using (BatchFilterForm form = new BatchFilterForm(batchFilterService, context))
            {
                form.ShowDialog();
            }
        }

        private void btnUnmergeAndFill_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int processedCount = unmergeAndFillService.Execute(Globals.ThisAddIn.Application);
                string message = processedCount == 0
                    ? "当前选区中没有合并单元格。"
                    : "已取消合并并填充 " + processedCount + " 个合并区域。";

                MessageBox.Show(message, "取消合并并填充", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    "操作失败：" + exception.Message,
                    "取消合并并填充",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                "eWorkHelper 是面向 Microsoft Excel 的工作效率插件。"
                + Environment.NewLine
                + Environment.NewLine
                + "版本：" + GetProductVersion(),
                "关于 eWorkHelper",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static string GetProductVersion()
        {
            Assembly assembly = typeof(MainRibbon).Assembly;
            AssemblyInformationalVersionAttribute attribute = (AssemblyInformationalVersionAttribute)Attribute.GetCustomAttribute(
                assembly,
                typeof(AssemblyInformationalVersionAttribute));

            return attribute == null || string.IsNullOrWhiteSpace(attribute.InformationalVersion)
                ? assembly.GetName().Version.ToString()
                : attribute.InformationalVersion;
        }
    }
}
