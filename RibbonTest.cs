using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace eWorkhelper
{
    public partial class MaimRibbon
    {
        private readonly BatchFilterService batchFilterService = new BatchFilterService();

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
    }
}
