namespace OutlookND
{
    partial class NoDelayRibbonVD : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public NoDelayRibbonVD()
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
            this.sendTab = this.Factory.CreateRibbonTab();
            this.sendGroup = this.Factory.CreateRibbonGroup();
            this.sendNoDelay = this.Factory.CreateRibbonButton();
            this.btn1m = this.Factory.CreateRibbonButton();
            this.btn5m = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.sendTab.SuspendLayout();
            this.sendGroup.SuspendLayout();
            // 
            // sendTab
            // 
            this.sendTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.sendTab.Groups.Add(this.sendGroup);
            this.sendTab.Label = "TIMED SEND";
            this.sendTab.Name = "sendTab";
            // 
            // sendGroup
            // 
            this.sendGroup.Items.Add(this.sendNoDelay);
            this.sendGroup.Items.Add(this.separator1);
            this.sendGroup.Items.Add(this.btn1m);
            this.sendGroup.Items.Add(this.btn5m);
            this.sendGroup.Label = "Send";
            this.sendGroup.Name = "sendGroup";
            // 
            // sendNoDelay
            // 
            this.sendNoDelay.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.sendNoDelay.Label = "Immediately";
            this.sendNoDelay.Name = "sendNoDelay";
            this.sendNoDelay.OfficeImageId = "DelayDeliveryOutlook";
            this.sendNoDelay.ShowImage = true;
            this.sendNoDelay.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sendNoDelay_Click);
            // 
            // btn1m
            // 
            this.btn1m.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn1m.Label = "In 1m";
            this.btn1m.Name = "btn1m";
            this.btn1m.OfficeImageId = "DelayDeliveryOutlook";
            this.btn1m.ShowImage = true;
            this.btn1m.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn1m_Click);
            // 
            // btn5m
            // 
            this.btn5m.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn5m.Label = "In 5m";
            this.btn5m.Name = "btn5m";
            this.btn5m.OfficeImageId = "DelayDeliveryOutlook";
            this.btn5m.ShowImage = true;
            this.btn5m.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn5m_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // NoDelayRibbonVD
            // 
            this.Name = "NoDelayRibbonVD";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.sendTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.NoDelayRibbonVD_Load);
            this.sendTab.ResumeLayout(false);
            this.sendTab.PerformLayout();
            this.sendGroup.ResumeLayout(false);
            this.sendGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab sendTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup sendGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton sendNoDelay;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn1m;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn5m;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal NoDelayRibbonVD NoDelayRibbonVD
        {
            get { return this.GetRibbon<NoDelayRibbonVD>(); }
        }
    }
}
