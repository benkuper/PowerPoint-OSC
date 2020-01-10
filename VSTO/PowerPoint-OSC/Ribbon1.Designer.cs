namespace PowerPoint_OSC
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.localPortInput = this.Factory.CreateRibbonEditBox();
            this.remoteHostInput = this.Factory.CreateRibbonEditBox();
            this.remotePortInput = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "OSC";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.localPortInput);
            this.group1.Items.Add(this.remoteHostInput);
            this.group1.Items.Add(this.remotePortInput);
            this.group1.Label = "Settings";
            this.group1.Name = "group1";
            // 
            // localPortInput
            // 
            this.localPortInput.Label = "Local Port";
            this.localPortInput.MaxLength = 5;
            this.localPortInput.Name = "localPortInput";
            this.localPortInput.Text = null;
            this.localPortInput.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.localPort_TextChanged);
            // 
            // remoteHostInput
            // 
            this.remoteHostInput.Label = "Remote Host";
            this.remoteHostInput.Name = "remoteHostInput";
            this.remoteHostInput.Text = null;
            // 
            // remotePortInput
            // 
            this.remotePortInput.Label = "Remote Port";
            this.remotePortInput.MaxLength = 5;
            this.remotePortInput.Name = "remotePortInput";
            this.remotePortInput.Text = null;
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox localPortInput;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox remoteHostInput;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox remotePortInput;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
