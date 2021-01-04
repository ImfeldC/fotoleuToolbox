namespace fotoleuToolbox
{
    partial class RibbonButton : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonButton()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonButton));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Label = this.Factory.CreateRibbonGroup();
            this.buttonDocument = this.Factory.CreateRibbonButton();
            this.buttonBill = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Label.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.Label);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // Label
            // 
            this.Label.Items.Add(this.buttonDocument);
            this.Label.Items.Add(this.buttonBill);
            this.Label.Label = "fotoleu Toolbox";
            this.Label.Name = "Label";
            // 
            // buttonDocument
            // 
            this.buttonDocument.Image = global::fotoleuToolbox.Properties.Resources.edit_document;
            this.buttonDocument.Label = "Generate Document";
            this.buttonDocument.Name = "buttonDocument";
            this.buttonDocument.ShowImage = true;
            this.buttonDocument.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonDocument_Click);
            // 
            // buttonBill
            // 
            this.buttonBill.Image = ((System.Drawing.Image)(resources.GetObject("buttonBill.Image")));
            this.buttonBill.Label = "Generate Bill";
            this.buttonBill.Name = "buttonBill";
            this.buttonBill.ShowImage = true;
            this.buttonBill.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGenerate_Click);
            // 
            // RibbonButton
            // 
            this.Name = "RibbonButton";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonButton_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Label.ResumeLayout(false);
            this.Label.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Label;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonBill;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonDocument;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonButton RibbonButton
        {
            get { return this.GetRibbon<RibbonButton>(); }
        }
    }
}
