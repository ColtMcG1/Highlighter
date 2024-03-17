namespace Highlighter
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.Highlighter = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.Toggle = this.Factory.CreateRibbonToggleButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.Color = this.Factory.CreateRibbonButton();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.Highlighter.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // Highlighter
            // 
            this.Highlighter.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.Highlighter.Groups.Add(this.group1);
            this.Highlighter.Groups.Add(this.group2);
            this.Highlighter.Label = "Highlighter";
            this.Highlighter.Name = "Highlighter";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Toggle);
            this.group1.Label = "Toggle";
            this.group1.Name = "group1";
            // 
            // Toggle
            // 
            this.Toggle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Toggle.Label = "Highlight";
            this.Toggle.Name = "Toggle";
            this.Toggle.ShowImage = true;
            // 
            // group2
            // 
            this.group2.Items.Add(this.Color);
            this.group2.Label = "Color";
            this.group2.Name = "group2";
            // 
            // Color
            // 
            this.Color.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Color.Label = "Color";
            this.Color.Name = "Color";
            this.Color.ShowImage = true;
            // 
            // colorDialog1
            // 
            this.colorDialog1.Color = System.Drawing.Color.IndianRed;
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.Highlighter);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.Highlighter.ResumeLayout(false);
            this.Highlighter.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab Highlighter;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton Toggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        private System.Windows.Forms.ColorDialog colorDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Color;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
