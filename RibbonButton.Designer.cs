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
            this.btnAuftragsblatt = this.Factory.CreateRibbonButton();
            this.btnQR = this.Factory.CreateRibbonButton();
            this.btnRechnung = this.Factory.CreateRibbonButton();
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
            this.Label.Items.Add(this.btnAuftragsblatt);
            this.Label.Items.Add(this.btnQR);
            this.Label.Items.Add(this.btnRechnung);
            this.Label.Label = "fotoleu Toolbox";
            this.Label.Name = "Label";
            // 
            // btnAuftragsblatt
            // 
            this.btnAuftragsblatt.Image = global::fotoleuToolbox.Properties.Resources.edit_document;
            this.btnAuftragsblatt.Label = "Erzeuge Auftragsblatt";
            this.btnAuftragsblatt.Name = "btnAuftragsblatt";
            this.btnAuftragsblatt.ShowImage = true;
            this.btnAuftragsblatt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAuftragsblatt_Click);
            // 
            // btnQR
            // 
            this.btnQR.Image = ((System.Drawing.Image)(resources.GetObject("btnQR.Image")));
            this.btnQR.Label = "Erzeuge QR EZS";
            this.btnQR.Name = "btnQR";
            this.btnQR.ShowImage = true;
            this.btnQR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnQR_Click);
            // 
            // btnRechnung
            // 
            this.btnRechnung.Image = global::fotoleuToolbox.Properties.Resources.multiple_files;
            this.btnRechnung.Label = "Erzeuge Rechnung";
            this.btnRechnung.Name = "btnRechnung";
            this.btnRechnung.ShowImage = true;
            this.btnRechnung.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRechnung_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnQR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAuftragsblatt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRechnung;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonButton RibbonButton
        {
            get { return this.GetRibbon<RibbonButton>(); }
        }
    }
}
