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
            // E-04：整段回调受保护；异常越过 VSTO Ribbon 边界会触发 CLR 未处理异常对话框，
            // 反复失败还可能导致加载项被 Excel 禁用。
            BatchFilterContext context = null;
            try
            {
                string errorMessage;
                if (!batchFilterService.TryCreateContext(Globals.ThisAddIn.Application, out context, out errorMessage))
                {
                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        MessageBox.Show(errorMessage, "批量过滤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    return;
                }

                // BatchFilterForm.Dispose 负责释放上下文持有的 RCW（E-03）。
                BatchFilterForm form = null;
                try
                {
                    form = new BatchFilterForm(batchFilterService, context);
                    context = null;
                    form.ShowDialog();
                }
                finally
                {
                    if (form != null)
                    {
                        form.Dispose();
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    "批量过滤失败：" + exception.Message,
                    "批量过滤",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (context != null)
                {
                    context.Dispose();
                }
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
