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
            this.txbSelectedRange = this.Factory.CreateRibbonEditBox();
            this.btnSetReplacementRange = this.Factory.CreateRibbonButton();
            this.btnReplace = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.txbTarget = this.Factory.CreateRibbonEditBox();
            this.txbReplacement = this.Factory.CreateRibbonEditBox();
            this.btnReplaceSubString = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.txbSelectedRange);
            this.group1.Items.Add(this.btnSetReplacementRange);
            this.group1.Items.Add(this.btnReplace);
            this.group1.Label = "Замена типа Ключ-Значение";
            this.group1.Name = "group1";
            // 
            // txbSelectedRange
            // 
            this.txbSelectedRange.Label = "Выбранный диапазон";
            this.txbSelectedRange.Name = "txbSelectedRange";
            this.txbSelectedRange.SizeString = "A12:A13";
            this.txbSelectedRange.Text = null;
            // 
            // btnSetReplacementRange
            // 
            this.btnSetReplacementRange.Label = "Дипазон замены";
            this.btnSetReplacementRange.Name = "btnSetReplacementRange";
            this.btnSetReplacementRange.ScreenTip = "Выберете диапазон из двух столбцов \"Что заменить\" - \"На что заменить\"";
            this.btnSetReplacementRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetReplacementRange_Click);
            // 
            // btnReplace
            // 
            this.btnReplace.Label = "Заменить";
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplace_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.txbTarget);
            this.group2.Items.Add(this.txbReplacement);
            this.group2.Items.Add(this.btnReplaceSubString);
            this.group2.Label = "Замена части строки";
            this.group2.Name = "group2";
            // 
            // txbTarget
            // 
            this.txbTarget.Label = "Заменить строку:";
            this.txbTarget.Name = "txbTarget";
            this.txbTarget.SizeString = "ПРЕОБРАЗОВ;;";
            this.txbTarget.Text = null;
            // 
            // txbReplacement
            // 
            this.txbReplacement.Label = "Заменить на:";
            this.txbReplacement.Name = "txbReplacement";
            this.txbReplacement.SizeString = "ПРЕОБРАЗОВАНИЕ";
            this.txbReplacement.Text = null;
            // 
            // btnReplaceSubString
            // 
            this.btnReplaceSubString.Label = "Заменить";
            this.btnReplaceSubString.Name = "btnReplaceSubString";
            this.btnReplaceSubString.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplaceSubString_Click);
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplace;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txbSelectedRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetReplacementRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txbTarget;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txbReplacement;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplaceSubString;
    }

    partial class ThisRibbonCollection
    {
        internal MainRibbon MainRibbon
        {
            get { return this.GetRibbon<MainRibbon>(); }
        }
    }
}
