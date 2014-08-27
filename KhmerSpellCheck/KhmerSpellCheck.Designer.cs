namespace KhmerSpellCheck
{
    partial class KhmerSpellCheck : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public KhmerSpellCheck()
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
            this.grpKhmerSpellCheck = this.Factory.CreateRibbonGroup();
            this.btnSpellCheck = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpKhmerSpellCheck.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpKhmerSpellCheck);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpKhmerSpellCheck
            // 
            this.grpKhmerSpellCheck.Items.Add(this.btnSpellCheck);
            this.grpKhmerSpellCheck.Label = "Khmer Proofing";
            this.grpKhmerSpellCheck.Name = "grpKhmerSpellCheck";
            // 
            // btnSpellCheck
            // 
            this.btnSpellCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSpellCheck.Image = global::KhmerSpellCheck.Properties.Resources.spellcheck;
            this.btnSpellCheck.Label = "SBBIC Khmer Spell Checker";
            this.btnSpellCheck.Name = "btnSpellCheck";
            this.btnSpellCheck.ScreenTip = "SBBIC Khmer Spell Checker";
            this.btnSpellCheck.ShowImage = true;
            this.btnSpellCheck.SuperTip = "Check the spelling of the Khmer text in the document";
            this.btnSpellCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSpellCheck_Click);
            // 
            // KhmerSpellCheck
            // 
            this.Name = "KhmerSpellCheck";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.KhmerSpellCheck_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpKhmerSpellCheck.ResumeLayout(false);
            this.grpKhmerSpellCheck.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpKhmerSpellCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSpellCheck;
    }

    partial class ThisRibbonCollection
    {
        internal KhmerSpellCheck KhmerSpellCheck
        {
            get { return this.GetRibbon<KhmerSpellCheck>(); }
        }
    }
}
