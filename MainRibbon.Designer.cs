namespace eWorkhelper
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainRibbon));
            this.tabEWorkHelper = this.Factory.CreateRibbonTab();
            this.grpDataTools = this.Factory.CreateRibbonGroup();
            this.btnBatchFilter = this.Factory.CreateRibbonButton();
            this.btnUnmergeAndFill = this.Factory.CreateRibbonButton();
            this.btnExcelDiff = this.Factory.CreateRibbonButton();
            this.grpAbout = this.Factory.CreateRibbonGroup();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabEWorkHelper.SuspendLayout();
            this.grpDataTools.SuspendLayout();
            this.grpAbout.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEWorkHelper
            // 
            this.tabEWorkHelper.Groups.Add(this.grpDataTools);
            this.tabEWorkHelper.Groups.Add(this.grpAbout);
            this.tabEWorkHelper.Label = "工作助手";
            this.tabEWorkHelper.Name = "tabEWorkHelper";
            // 
            // grpDataTools
            // 
            this.grpDataTools.Items.Add(this.btnBatchFilter);
            this.grpDataTools.Items.Add(this.btnUnmergeAndFill);
            this.grpDataTools.Items.Add(this.btnExcelDiff);
            this.grpDataTools.Label = "数据工具";
            this.grpDataTools.Name = "grpDataTools";
            // 
            // btnBatchFilter
            // 
            this.btnBatchFilter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnBatchFilter.Image = ((System.Drawing.Image)(resources.GetObject("btnBatchFilter.Image")));
            this.btnBatchFilter.Label = "批量过滤";
            this.btnBatchFilter.Name = "btnBatchFilter";
            this.btnBatchFilter.ShowImage = true;
            this.btnBatchFilter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBatchFilter_Click);
            // 
            // btnUnmergeAndFill
            // 
            this.btnUnmergeAndFill.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUnmergeAndFill.Image = global::eWorkhelper.Properties.Resources.mergecell;
            this.btnUnmergeAndFill.Label = "取消合并";
            this.btnUnmergeAndFill.Name = "btnUnmergeAndFill";
            this.btnUnmergeAndFill.OfficeImageId = "UnmergeCells";
            this.btnUnmergeAndFill.ShowImage = true;
            this.btnUnmergeAndFill.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUnmergeAndFill_Click);
            // 
            // btnExcelDiff
            // 
            this.btnExcelDiff.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExcelDiff.Image = global::eWorkhelper.Properties.Resources.Diff;
            this.btnExcelDiff.Label = "差异对比";
            this.btnExcelDiff.Name = "btnExcelDiff";
            this.btnExcelDiff.ShowImage = true;
            this.btnExcelDiff.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExcelDiff_Click);
            // 
            // grpAbout
            // 
            this.grpAbout.Items.Add(this.btnAbout);
            this.grpAbout.Label = "关于";
            this.grpAbout.Name = "grpAbout";
            // 
            // btnAbout
            // 
            this.btnAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAbout.Image = global::eWorkhelper.Properties.Resources.info;
            this.btnAbout.Label = "关于";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.ShowImage = true;
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabEWorkHelper);
            this.tabEWorkHelper.ResumeLayout(false);
            this.tabEWorkHelper.PerformLayout();
            this.grpDataTools.ResumeLayout(false);
            this.grpDataTools.PerformLayout();
            this.grpAbout.ResumeLayout(false);
            this.grpAbout.PerformLayout();
            this.ResumeLayout(false);

        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEWorkHelper;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpDataTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBatchFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUnmergeAndFill;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExcelDiff;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
    }
}
