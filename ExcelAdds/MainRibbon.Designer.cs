namespace ExcelAdds
{
    partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MainRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.txbTarget = this.Factory.CreateRibbonEditBox();
            this.txbNewText = this.Factory.CreateRibbonEditBox();
            this.btnReplace = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.txbSelectedRange = this.Factory.CreateRibbonEditBox();
            this.btnSetReplacementRange = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.txbTarget);
            this.group1.Items.Add(this.txbNewText);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.txbSelectedRange);
            this.group1.Items.Add(this.btnSetReplacementRange);
            this.group1.Items.Add(this.btnReplace);
            this.group1.Label = "Замена части строки";
            this.group1.Name = "group1";
            // 
            // txbTarget
            // 
            this.txbTarget.Label = "Заменить строку:";
            this.txbTarget.Name = "txbTarget";
            // 
            // txbNewText
            // 
            this.txbNewText.Label = "Заменить на:";
            this.txbNewText.Name = "txbNewText";
            // 
            // btnReplace
            // 
            this.btnReplace.Label = "Заменить";
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplace_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // txbSelectedRange
            // 
            this.txbSelectedRange.Label = "Выбранный диапазон";
            this.txbSelectedRange.Name = "txbSelectedRange";
            // 
            // btnSetReplacementRange
            // 
            this.btnSetReplacementRange.Label = "Дипазон замены";
            this.btnSetReplacementRange.Name = "btnSetReplacementRange";
            this.btnSetReplacementRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetReplacementRange_Click);
            // 
            // MainRibbon
            // 
            this.Name = "MainRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MainRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txbTarget;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txbNewText;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplace;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txbSelectedRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetReplacementRange;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
