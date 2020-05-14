namespace OpenXmlForVsto.Excel.BenchmarkAddin {
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.BenchmarkAddinTab = this.Factory.CreateRibbonTab();
            this.GroupVsto = this.Factory.CreateRibbonGroup();
            this.ButtonWrite1CellVsto = this.Factory.CreateRibbonButton();
            this.ButtonWrite100CellsVsto = this.Factory.CreateRibbonButton();
            this.ButtonWrite10kCellsVsto = this.Factory.CreateRibbonButton();
            this.ButtonWrite1mCellsVsto = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.ButtonRead1CellVsto = this.Factory.CreateRibbonButton();
            this.ButtonRead100CellsVsto = this.Factory.CreateRibbonButton();
            this.ButtonRead10kCellsVsto = this.Factory.CreateRibbonButton();
            this.ButtonRead1mCellsVsto = this.Factory.CreateRibbonButton();
            this.GroupOpenXML = this.Factory.CreateRibbonGroup();
            this.ButtonWrite1CellOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonWrite100CellsOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonWrite10kCellsOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonWrite1mCellsOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonWrite9mCellsOpenXML = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.ButtonRead1CellOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonRead100CellsOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonRead10kCellsOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonRead1mCellsOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonRead9mCellsOpenXML = this.Factory.CreateRibbonButton();
            this.BenchmarkAddinTab.SuspendLayout();
            this.GroupVsto.SuspendLayout();
            this.GroupOpenXML.SuspendLayout();
            this.SuspendLayout();
            // 
            // BenchmarkAddinTab
            // 
            this.BenchmarkAddinTab.Groups.Add(this.GroupVsto);
            this.BenchmarkAddinTab.Groups.Add(this.GroupOpenXML);
            this.BenchmarkAddinTab.Label = "Benchmark";
            this.BenchmarkAddinTab.Name = "BenchmarkAddinTab";
            // 
            // GroupVsto
            // 
            this.GroupVsto.Items.Add(this.ButtonWrite1CellVsto);
            this.GroupVsto.Items.Add(this.ButtonWrite100CellsVsto);
            this.GroupVsto.Items.Add(this.ButtonWrite10kCellsVsto);
            this.GroupVsto.Items.Add(this.ButtonWrite1mCellsVsto);
            this.GroupVsto.Items.Add(this.separator1);
            this.GroupVsto.Items.Add(this.ButtonRead1CellVsto);
            this.GroupVsto.Items.Add(this.ButtonRead100CellsVsto);
            this.GroupVsto.Items.Add(this.ButtonRead10kCellsVsto);
            this.GroupVsto.Items.Add(this.ButtonRead1mCellsVsto);
            this.GroupVsto.Label = "VSTO";
            this.GroupVsto.Name = "GroupVsto";
            // 
            // ButtonWrite1CellVsto
            // 
            this.ButtonWrite1CellVsto.Label = "Write 1 cell";
            this.ButtonWrite1CellVsto.Name = "ButtonWrite1CellVsto";
            this.ButtonWrite1CellVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite1CellVsto_Click);
            // 
            // ButtonWrite100CellsVsto
            // 
            this.ButtonWrite100CellsVsto.Label = "Write 100 cells";
            this.ButtonWrite100CellsVsto.Name = "ButtonWrite100CellsVsto";
            this.ButtonWrite100CellsVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite100CellsVsto_Click);
            // 
            // ButtonWrite10kCellsVsto
            // 
            this.ButtonWrite10kCellsVsto.Label = "Write 10k cells";
            this.ButtonWrite10kCellsVsto.Name = "ButtonWrite10kCellsVsto";
            this.ButtonWrite10kCellsVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite10kCellsVsto_Click);
            // 
            // ButtonWrite1mCellsVsto
            // 
            this.ButtonWrite1mCellsVsto.Label = "Write 1m cells";
            this.ButtonWrite1mCellsVsto.Name = "ButtonWrite1mCellsVsto";
            this.ButtonWrite1mCellsVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite1mCellsVsto_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // ButtonRead1CellVsto
            // 
            this.ButtonRead1CellVsto.Label = "Read 1 cell";
            this.ButtonRead1CellVsto.Name = "ButtonRead1CellVsto";
            this.ButtonRead1CellVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRead1CellVsto_Click);
            // 
            // ButtonRead100CellsVsto
            // 
            this.ButtonRead100CellsVsto.Label = "Read 100 cells";
            this.ButtonRead100CellsVsto.Name = "ButtonRead100CellsVsto";
            this.ButtonRead100CellsVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRead100CellsVsto_Click);
            // 
            // ButtonRead10kCellsVsto
            // 
            this.ButtonRead10kCellsVsto.Label = "Read 10k cells";
            this.ButtonRead10kCellsVsto.Name = "ButtonRead10kCellsVsto";
            this.ButtonRead10kCellsVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRead10kCellsVsto_Click);
            // 
            // ButtonRead1mCellsVsto
            // 
            this.ButtonRead1mCellsVsto.Label = "Read 1m cells";
            this.ButtonRead1mCellsVsto.Name = "ButtonRead1mCellsVsto";
            this.ButtonRead1mCellsVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRead1mCellsVsto_Click);
            // 
            // GroupOpenXML
            // 
            this.GroupOpenXML.Items.Add(this.ButtonWrite1CellOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonWrite100CellsOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonWrite10kCellsOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonWrite1mCellsOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonWrite9mCellsOpenXML);
            this.GroupOpenXML.Items.Add(this.separator3);
            this.GroupOpenXML.Items.Add(this.ButtonRead1CellOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonRead100CellsOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonRead10kCellsOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonRead1mCellsOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonRead9mCellsOpenXML);
            this.GroupOpenXML.Label = "OpenXML";
            this.GroupOpenXML.Name = "GroupOpenXML";
            // 
            // ButtonWrite1CellOpenXML
            // 
            this.ButtonWrite1CellOpenXML.Label = "Write 1 cell";
            this.ButtonWrite1CellOpenXML.Name = "ButtonWrite1CellOpenXML";
            this.ButtonWrite1CellOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite1CellOpenXML_Click);
            // 
            // ButtonWrite100CellsOpenXML
            // 
            this.ButtonWrite100CellsOpenXML.Label = "Write 100 cells";
            this.ButtonWrite100CellsOpenXML.Name = "ButtonWrite100CellsOpenXML";
            this.ButtonWrite100CellsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite100CellsOpenXML_Click);
            // 
            // ButtonWrite10kCellsOpenXML
            // 
            this.ButtonWrite10kCellsOpenXML.Label = "Write 10k cells";
            this.ButtonWrite10kCellsOpenXML.Name = "ButtonWrite10kCellsOpenXML";
            this.ButtonWrite10kCellsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite10kCellsOpenXML_Click);
            // 
            // ButtonWrite1mCellsOpenXML
            // 
            this.ButtonWrite1mCellsOpenXML.Label = "Write 1m cells";
            this.ButtonWrite1mCellsOpenXML.Name = "ButtonWrite1mCellsOpenXML";
            this.ButtonWrite1mCellsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite1mCellsOpenXML_Click);
            // 
            // ButtonWrite9mCellsOpenXML
            // 
            this.ButtonWrite9mCellsOpenXML.Label = "Write 9m cells";
            this.ButtonWrite9mCellsOpenXML.Name = "ButtonWrite9mCellsOpenXML";
            this.ButtonWrite9mCellsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite9mCellsOpenXML_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // ButtonRead1CellOpenXML
            // 
            this.ButtonRead1CellOpenXML.Label = "Read 1 cell";
            this.ButtonRead1CellOpenXML.Name = "ButtonRead1CellOpenXML";
            this.ButtonRead1CellOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRead1CellOpenXML_Click);
            // 
            // ButtonRead100CellsOpenXML
            // 
            this.ButtonRead100CellsOpenXML.Label = "Read 100 cells";
            this.ButtonRead100CellsOpenXML.Name = "ButtonRead100CellsOpenXML";
            this.ButtonRead100CellsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRead100CellsOpenXML_Click);
            // 
            // ButtonRead10kCellsOpenXML
            // 
            this.ButtonRead10kCellsOpenXML.Label = "Read 10k cells";
            this.ButtonRead10kCellsOpenXML.Name = "ButtonRead10kCellsOpenXML";
            this.ButtonRead10kCellsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRead10kCellsOpenXML_Click);
            // 
            // ButtonRead1mCellsOpenXML
            // 
            this.ButtonRead1mCellsOpenXML.Label = "Read 1m cells";
            this.ButtonRead1mCellsOpenXML.Name = "ButtonRead1mCellsOpenXML";
            this.ButtonRead1mCellsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRead1mCellsOpenXML_Click);
            // 
            // ButtonRead9mCellsOpenXML
            // 
            this.ButtonRead9mCellsOpenXML.Label = "Read 9m cells";
            this.ButtonRead9mCellsOpenXML.Name = "ButtonRead9mCellsOpenXML";
            this.ButtonRead9mCellsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonRead9mCellsOpenXML_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.BenchmarkAddinTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.BenchmarkAddinTab.ResumeLayout(false);
            this.BenchmarkAddinTab.PerformLayout();
            this.GroupVsto.ResumeLayout(false);
            this.GroupVsto.PerformLayout();
            this.GroupOpenXML.ResumeLayout(false);
            this.GroupOpenXML.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab BenchmarkAddinTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonRead1CellVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonRead100CellsVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonRead10kCellsVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonRead1mCellsVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite1CellVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite100CellsVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite10kCellsVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite1mCellsVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonRead1CellOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonRead100CellsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonRead10kCellsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonRead1mCellsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonRead9mCellsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite1CellOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite100CellsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite10kCellsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite1mCellsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite9mCellsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
    }

    partial class ThisRibbonCollection {
        internal Ribbon1 Ribbon1 {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
