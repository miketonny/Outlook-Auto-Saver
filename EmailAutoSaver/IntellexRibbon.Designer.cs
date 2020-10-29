namespace EmailAutoSaver
{
    partial class IntellexRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public IntellexRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(IntellexRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnAddProject = this.Factory.CreateRibbonButton();
            this.btnAddArchive = this.Factory.CreateRibbonButton();
            this.btnReLoad = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "INTELLEX";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnAddProject);
            this.group1.Items.Add(this.btnAddArchive);
            this.group1.Items.Add(this.btnReLoad);
            this.group1.Label = "Emails";
            this.group1.Name = "group1";
            // 
            // btnAddProject
            // 
            this.btnAddProject.Image = ((System.Drawing.Image)(resources.GetObject("btnAddProject.Image")));
            this.btnAddProject.Label = "Add New/Existing Project";
            this.btnAddProject.Name = "btnAddProject";
            this.btnAddProject.ShowImage = true;
            this.btnAddProject.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddProject_Click);
            // 
            // btnAddArchive
            // 
            this.btnAddArchive.Image = ((System.Drawing.Image)(resources.GetObject("btnAddArchive.Image")));
            this.btnAddArchive.Label = "Add Archived Project";
            this.btnAddArchive.Name = "btnAddArchive";
            this.btnAddArchive.ShowImage = true;
            this.btnAddArchive.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddArchive_Click);
            // 
            // btnReLoad
            // 
            this.btnReLoad.Image = ((System.Drawing.Image)(resources.GetObject("btnReLoad.Image")));
            this.btnReLoad.Label = "Refresh Hooks";
            this.btnReLoad.Name = "btnReLoad";
            this.btnReLoad.ShowImage = true;
            this.btnReLoad.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReLoad_Click);
            // 
            // IntellexRibbon
            // 
            this.Name = "IntellexRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.IntellexRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddProject;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReLoad;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddArchive;
    }

    partial class ThisRibbonCollection
    {
        internal IntellexRibbon IntellexRibbon
        {
            get { return this.GetRibbon<IntellexRibbon>(); }
        }
    }
}
