namespace eWorkhelper
{
    partial class MaimRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public MaimRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tabEWorkHelper = this.Factory.CreateRibbonTab();
            this.grpDataTools = this.Factory.CreateRibbonGroup();
            this.btnBatchFilter = this.Factory.CreateRibbonButton();
            this.tabEWorkHelper.SuspendLayout();
            this.grpDataTools.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEWorkHelper
            // 
            this.tabEWorkHelper.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEWorkHelper.Groups.Add(this.grpDataTools);
            this.tabEWorkHelper.Label = "eWorkHelper";
            this.tabEWorkHelper.Name = "tabEWorkHelper";
            // 
            // grpDataTools
            // 
            this.grpDataTools.Items.Add(this.btnBatchFilter);
            this.grpDataTools.Label = "数据工具";
            this.grpDataTools.Name = "grpDataTools";
            // 
            // btnBatchFilter
            // 
            this.btnBatchFilter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnBatchFilter.Label = "批量过滤";
            this.btnBatchFilter.Name = "btnBatchFilter";
            this.btnBatchFilter.ShowImage = true;
            this.btnBatchFilter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBatchFilter_Click);
            // 
            // MaimRibbon
            // 
            this.Name = "MaimRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEWorkHelper);
            this.tabEWorkHelper.ResumeLayout(false);
            this.tabEWorkHelper.PerformLayout();
            this.grpDataTools.ResumeLayout(false);
            this.grpDataTools.PerformLayout();
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEWorkHelper;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDataTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBatchFilter;
    }
}
