namespace SDV_Outlook_Tools
{
    partial class rb_sdv_outlook_tools : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public rb_sdv_outlook_tools()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(rb_sdv_outlook_tools));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_RemoveAttachments = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "SDV Outlook Tools";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_RemoveAttachments);
            this.group1.Name = "group1";
            // 
            // btn_RemoveAttachments
            // 
            this.btn_RemoveAttachments.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_RemoveAttachments.Image = ((System.Drawing.Image)(resources.GetObject("btn_RemoveAttachments.Image")));
            this.btn_RemoveAttachments.Label = "Remove Attachments";
            this.btn_RemoveAttachments.Name = "btn_RemoveAttachments";
            this.btn_RemoveAttachments.ScreenTip = "entfernt alle Mailanhänge von allen gelesenen E-Mails älter 90 Tage.";
            this.btn_RemoveAttachments.ShowImage = true;
            this.btn_RemoveAttachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_RemoveAttachments_Click);
            // 
            // rb_sdv_outlook_tools
            // 
            this.Name = "rb_sdv_outlook_tools";
            this.RibbonType = "Microsoft.Outlook.Explorer";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_RemoveAttachments;
    }

    partial class ThisRibbonCollection
    {
        internal rb_sdv_outlook_tools Ribbon1
        {
            get { return this.GetRibbon<rb_sdv_outlook_tools>(); }
        }
    }
}
