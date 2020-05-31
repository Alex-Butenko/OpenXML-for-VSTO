namespace OpenXmlForVsto.Word.BenchmarkAddin {
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
            this.ButtonWrite1RunVsto = this.Factory.CreateRibbonButton();
            this.ButtonWrite100RunsVsto = this.Factory.CreateRibbonButton();
            this.ButtonWrite10kRunsVsto = this.Factory.CreateRibbonButton();
            this.ButtonWrite100kRunsVsto = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.ButtonReadVsto = this.Factory.CreateRibbonButton();
            this.GroupOpenXML = this.Factory.CreateRibbonGroup();
            this.ButtonWrite1RunOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonWrite100RunsOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonWrite10kRunsOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonWrite100kRunsOpenXML = this.Factory.CreateRibbonButton();
            this.ButtonWrite1mRunsOpenXML = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.ButtonReadOpenXML = this.Factory.CreateRibbonButton();
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
            this.GroupVsto.Items.Add(this.ButtonWrite1RunVsto);
            this.GroupVsto.Items.Add(this.ButtonWrite100RunsVsto);
            this.GroupVsto.Items.Add(this.ButtonWrite10kRunsVsto);
            this.GroupVsto.Items.Add(this.ButtonWrite100kRunsVsto);
            this.GroupVsto.Items.Add(this.separator1);
            this.GroupVsto.Items.Add(this.ButtonReadVsto);
            this.GroupVsto.Label = "VSTO";
            this.GroupVsto.Name = "GroupVsto";
            // 
            // ButtonWrite1RunVsto
            // 
            this.ButtonWrite1RunVsto.Label = "Write 1 run";
            this.ButtonWrite1RunVsto.Name = "ButtonWrite1RunVsto";
            this.ButtonWrite1RunVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite1RunVsto_Click);
            // 
            // ButtonWrite100RunsVsto
            // 
            this.ButtonWrite100RunsVsto.Label = "Write 100 runs";
            this.ButtonWrite100RunsVsto.Name = "ButtonWrite100RunsVsto";
            this.ButtonWrite100RunsVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite100RunsVsto_Click);
            // 
            // ButtonWrite10kRunsVsto
            // 
            this.ButtonWrite10kRunsVsto.Label = "Write 10k runs";
            this.ButtonWrite10kRunsVsto.Name = "ButtonWrite10kRunsVsto";
            this.ButtonWrite10kRunsVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite10kRunsVsto_Click);
            // 
            // ButtonWrite100kRunsVsto
            // 
            this.ButtonWrite100kRunsVsto.Label = "Write 100k runs";
            this.ButtonWrite100kRunsVsto.Name = "ButtonWrite100kRunsVsto";
            this.ButtonWrite100kRunsVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite100kRunsVsto_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // ButtonReadVsto
            // 
            this.ButtonReadVsto.Label = "Read";
            this.ButtonReadVsto.Name = "ButtonReadVsto";
            this.ButtonReadVsto.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonReadVsto_Click);
            // 
            // GroupOpenXML
            // 
            this.GroupOpenXML.Items.Add(this.ButtonWrite1RunOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonWrite100RunsOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonWrite10kRunsOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonWrite100kRunsOpenXML);
            this.GroupOpenXML.Items.Add(this.ButtonWrite1mRunsOpenXML);
            this.GroupOpenXML.Items.Add(this.separator3);
            this.GroupOpenXML.Items.Add(this.ButtonReadOpenXML);
            this.GroupOpenXML.Label = "OpenXML";
            this.GroupOpenXML.Name = "GroupOpenXML";
            // 
            // ButtonWrite1RunOpenXML
            // 
            this.ButtonWrite1RunOpenXML.Label = "Write 1 run";
            this.ButtonWrite1RunOpenXML.Name = "ButtonWrite1RunOpenXML";
            this.ButtonWrite1RunOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite1RunOpenXML_Click);
            // 
            // ButtonWrite100RunsOpenXML
            // 
            this.ButtonWrite100RunsOpenXML.Label = "Write 100 runs";
            this.ButtonWrite100RunsOpenXML.Name = "ButtonWrite100RunsOpenXML";
            this.ButtonWrite100RunsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite100RunsOpenXML_Click);
            // 
            // ButtonWrite10kRunsOpenXML
            // 
            this.ButtonWrite10kRunsOpenXML.Label = "Write 10k runs";
            this.ButtonWrite10kRunsOpenXML.Name = "ButtonWrite10kRunsOpenXML";
            this.ButtonWrite10kRunsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite10kRunsOpenXML_Click);
            // 
            // ButtonWrite100kRunsOpenXML
            // 
            this.ButtonWrite100kRunsOpenXML.Label = "Write 100k runs";
            this.ButtonWrite100kRunsOpenXML.Name = "ButtonWrite100kRunsOpenXML";
            this.ButtonWrite100kRunsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite100kRunsOpenXML_Click);
            // 
            // ButtonWrite1mRunsOpenXML
            // 
            this.ButtonWrite1mRunsOpenXML.Label = "Write 1m runs";
            this.ButtonWrite1mRunsOpenXML.Name = "ButtonWrite1mRunsOpenXML";
            this.ButtonWrite1mRunsOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonWrite1mRunsOpenXML_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // ButtonReadOpenXML
            // 
            this.ButtonReadOpenXML.Label = "Read";
            this.ButtonReadOpenXML.Name = "ButtonReadOpenXML";
            this.ButtonReadOpenXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonReadOpenXML_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonReadVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite1RunVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite100RunsVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite10kRunsVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite100kRunsVsto;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonReadOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite1RunOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite100RunsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite10kRunsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite100kRunsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ButtonWrite1mRunsOpenXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
    }

    partial class ThisRibbonCollection {
        internal Ribbon1 Ribbon1 {
            get {
                return this.GetRibbon<Ribbon1>();
            }
        }
    }
}