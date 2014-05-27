using OutlookInboxCleaner;

namespace OutlookInboxCleaner
{
    partial class InboxCleanerRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public InboxCleanerRibbon()
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
            this.tabCleanup = this.Factory.CreateRibbonTab();
            this.cleaupGroup = this.Factory.CreateRibbonGroup();
            this.btnCleanup = this.Factory.CreateRibbonButton();
            this.btnStop = this.Factory.CreateRibbonButton();
            this.batchSize = this.Factory.CreateRibbonEditBox();
            this.deletePermanently = this.Factory.CreateRibbonCheckBox();
            this.tabCleanup.SuspendLayout();
            this.cleaupGroup.SuspendLayout();
            // 
            // tabCleanup
            // 
            this.tabCleanup.Groups.Add(this.cleaupGroup);
            this.tabCleanup.Label = "Cleanup";
            this.tabCleanup.Name = "tabCleanup";
            // 
            // cleaupGroup
            // 
            this.cleaupGroup.Items.Add(this.btnCleanup);
            this.cleaupGroup.Items.Add(this.btnStop);
            this.cleaupGroup.Items.Add(this.batchSize);
            this.cleaupGroup.Items.Add(this.deletePermanently);
            this.cleaupGroup.Label = "Cleanup";
            this.cleaupGroup.Name = "cleaupGroup";
            // 
            // btnCleanup
            // 
            this.btnCleanup.Label = "Cleanup Selected Folder";
            this.btnCleanup.Name = "btnCleanup";
            this.btnCleanup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanup_Click);
            // 
            // btnStop
            // 
            this.btnStop.Enabled = false;
            this.btnStop.Label = "Stop";
            this.btnStop.Name = "btnStop";
            this.btnStop.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStop_Click);
            // 
            // batchSize
            // 
            this.batchSize.Label = "Batch Size";
            this.batchSize.Name = "batchSize";
            this.batchSize.Text = "10";
            // 
            // deletePermanently
            // 
            this.deletePermanently.Label = "Delete Permanently";
            this.deletePermanently.Name = "deletePermanently";
            // 
            // InboxCleanerRibbon
            // 
            this.Name = "InboxCleanerRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabCleanup);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.InboxCleanerRibbon_Load);
            this.tabCleanup.ResumeLayout(false);
            this.tabCleanup.PerformLayout();
            this.cleaupGroup.ResumeLayout(false);
            this.cleaupGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabCleanup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup cleaupGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox batchSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStop;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox deletePermanently;
    }

    partial class ThisRibbonCollection
    {
        internal InboxCleanerRibbon InboxCleanerRibbon
        {
            get { return this.GetRibbon<InboxCleanerRibbon>(); }
        }
    }
}
