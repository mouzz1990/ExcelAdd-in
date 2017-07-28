using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace ExcelAdds
{
    public partial class ThisAddIn
    {
        //глобальные переменные
        Excel.Worksheet wSheet;
        Excel.Range SelectedTarget;
        IMainRibbonView ribbon;
        Dictionary<string, string> ReplacementDictionary;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            wSheet = Application.ActiveWorkbook.ActiveSheet;
            wSheet.SelectionChange += new Excel.DocEvents_SelectionChangeEventHandler(wSheet_SelectionChange);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            wSheet.SelectionChange -= new Excel.DocEvents_SelectionChangeEventHandler(wSheet_SelectionChange);
        }

        //создание ленты
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new MainRibbon();
            ribbon.ButtonReplaceClicked += ribbon_ButtonReplaceClicked;
            ribbon.ButtonSetReplacementRangeClicked += ribbon_ButtonSetReplacementRangeClicked;
            ribbon.ButtonReplaceSubStringClicked += ribbon_ButtonReplaceSubStringClicked;
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { (IRibbonExtension)ribbon });
        }

        //Замена строк
        void ribbon_ButtonReplaceSubStringClicked()
        {
            foreach (Excel.Range c in SelectedTarget.Cells)
            {
                string temp = c.Value;

                if (temp.Contains(ribbon.TargetString))
                    c.Value = ReplaceString(temp, ribbon.TargetString, ribbon.Replacement);
            }
        }

        //создание словаря для замены
        void ribbon_ButtonSetReplacementRangeClicked()
        {
            ribbon.SelectedRange = SelectedTarget.Address.Replace("$","");

            ReplacementDictionary = new Dictionary<string, string>();

            foreach (Excel.Range row in SelectedTarget.Rows)
            {
                ReplacementDictionary.Add(row.Cells[1].Value, row.Cells[2].Value);
            }
        }

        //Замена указанного диапазона созданными значениями словаря
        private void ribbon_ButtonReplaceClicked()
        {
            foreach (Excel.Range c in SelectedTarget.Cells)
            {
                foreach (var key in ReplacementDictionary.Keys)
                {
                    if (c.Value == key)
                        c.Value = ReplacementDictionary[key];
                }
            }
        }

        //метод замены строки
        string ReplaceString(string input, string find, string replace)
        {
            int startIndex = input.IndexOf(find);
            int lastIndex = startIndex + find.Length;

            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < startIndex; i++)
                sb.Append(input[i]);

            sb.Append(replace);

            for (int i = lastIndex; i < input.Length; i++)
                sb.Append(input[i]);

            return sb.ToString();
        }

        //получение выделенного рэнжа
        void wSheet_SelectionChange(Excel.Range Target)
        {
            SelectedTarget = Target;
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
