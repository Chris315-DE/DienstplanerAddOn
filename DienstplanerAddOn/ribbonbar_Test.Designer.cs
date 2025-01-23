namespace DienstplanerAddOn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.Setup = this.Factory.CreateRibbonGroup();
            this.btnTest = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btnLoadMitarbeiter = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Setup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Setup);
            this.tab1.Label = "DienstplanAddOn";
            this.tab1.Name = "tab1";
            // 
            // Setup
            // 
            this.Setup.Items.Add(this.btnTest);
            this.Setup.Items.Add(this.button1);
            this.Setup.Items.Add(this.btnLoadMitarbeiter);
            this.Setup.Label = "Setup";
            this.Setup.Name = "Setup";
            // 
            // btnTest
            // 
            this.btnTest.Label = "Setup Teil 1";
            this.btnTest.Name = "btnTest";
            this.btnTest.ShowImage = true;
            this.btnTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Label = "";
            this.button1.Name = "button1";
            // 
            // btnLoadMitarbeiter
            // 
            this.btnLoadMitarbeiter.Label = "Setup Teil 2";
            this.btnLoadMitarbeiter.Name = "btnLoadMitarbeiter";
            this.btnLoadMitarbeiter.ShowImage = true;
            this.btnLoadMitarbeiter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadMitarbeiter_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Setup.ResumeLayout(false);
            this.Setup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Setup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadMitarbeiter;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
