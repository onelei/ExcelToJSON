namespace ExcelToJSonAddIn
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.Key_R = this.Factory.CreateRibbonEditBox();
            this.Key_C = this.Factory.CreateRibbonEditBox();
            this.Button_XML = this.Factory.CreateRibbonButton();
            this.Value_R = this.Factory.CreateRibbonEditBox();
            this.Value_C = this.Factory.CreateRibbonEditBox();
            this.button_JSon = this.Factory.CreateRibbonButton();
            this.button_About = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Key_R);
            this.group1.Items.Add(this.Key_C);
            this.group1.Items.Add(this.button_JSon);
            this.group1.Items.Add(this.Value_R);
            this.group1.Items.Add(this.Value_C);
            this.group1.Items.Add(this.button_About);
            this.group1.Items.Add(this.Button_XML);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // Key_R
            // 
            this.Key_R.Label = "Key(行R)";
            this.Key_R.Name = "Key_R";
            this.Key_R.Text = null;
            // 
            // Key_C
            // 
            this.Key_C.Label = "Key(列C)";
            this.Key_C.Name = "Key_C";
            this.Key_C.Text = null;
            // 
            // Button_XML
            // 
            this.Button_XML.Label = "";
            this.Button_XML.Name = "Button_XML";
            // 
            // Value_R
            // 
            this.Value_R.Label = "Value(行R)";
            this.Value_R.Name = "Value_R";
            this.Value_R.Text = null;
            // 
            // Value_C
            // 
            this.Value_C.Label = "Value(列C)";
            this.Value_C.Name = "Value_C";
            this.Value_C.Text = null;
            // 
            // button_JSon
            // 
            this.button_JSon.Label = "ExcelToJSON";
            this.button_JSon.Name = "button_JSon";
            this.button_JSon.ShowImage = true;
            this.button_JSon.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_JSon_Click);
            // 
            // button_About
            // 
            this.button_About.Label = "About";
            this.button_About.Name = "button_About";
            this.button_About.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_About_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_JSon;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Key_R;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_About;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Key_C;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Value_R;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox Value_C;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Button_XML;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
