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
        Excel.Application eApp;
        Excel.Worksheet wSheet;
        Excel.Range SelectedTarget;
        IMainRibbonView ribbon;
        Dictionary<string, string> ReplacementDictionary;

        //Запуск
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            eApp = Application;
            eApp.ActiveWorkbook.SheetSelectionChange += ActiveWorkbook_SheetSelectionChange;
        }

        //Остановка
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            
        }

        //Выделение ячеек и запись ренжа согласно выделенного листа
        private void ActiveWorkbook_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            eApp.ActiveWorkbook.SheetSelectionChange -= ActiveWorkbook_SheetSelectionChange;
            wSheet = (Excel.Worksheet)Sh;
            SelectedTarget = Target;
            eApp.ActiveWorkbook.SheetSelectionChange += ActiveWorkbook_SheetSelectionChange;
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
            int replaceCounter = 0;
            int resetCounter = 0;

            foreach (Excel.Range c in SelectedTarget.Cells)
            {
                string temp = c.Value;

                if (temp == null)
                {
                    resetCounter++;

                    if (resetCounter >= 1000) break;

                    continue;
                }

                if (temp.Contains(ribbon.TargetString))
                {
                    resetCounter = 0;
                    replaceCounter++;

                    c.Value = ReplaceString(temp, ribbon.TargetString, ribbon.Replacement);
                }
            }
            MessageBox.Show(string.Format("Операция успешна!{0}{0}Произведено замен: {1}",
                            Environment.NewLine,
                            replaceCounter
                            ),
                            "Замена завершена успешно!",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                            );
        }

        //создание словаря для замены из выбранного диапазона 2-х столбцов
        void ribbon_ButtonSetReplacementRangeClicked()
        {
            ribbon.SelectedRange = SelectedTarget.Address.Replace("$", "");

            ReplacementDictionary = new Dictionary<string, string>();

            foreach (Excel.Range row in SelectedTarget.Rows)
            {
                if (row.Cells[1].Value == null) continue;
                if (ReplacementDictionary.ContainsKey(row.Cells[1].Value)) continue;

                ReplacementDictionary.Add(row.Cells[1].Value, row.Cells[2].Value);
            }

        }

        //Замена указанного диапазона созданными значениями словаря
        private void ribbon_ButtonReplaceClicked()
        {
            int resetCounter = 0;
            int counterReplace = 0;

            foreach (Excel.Range c in SelectedTarget.Cells)
            {
                if (c.Value == null)
                {
                    resetCounter++;

                    if (resetCounter >= 1000)  break;

                    continue;
                }

                foreach (var key in ReplacementDictionary.Keys)
                {
                    resetCounter = 0;

                    if (c.Value == key)
                    {
                        c.Value = ReplacementDictionary[key];
                        counterReplace++;
                    }
                }
            }

            MessageBox.Show(string.Format("Операция успешна!{0}{0}Произведено замен: {1}",
                            Environment.NewLine,
                            counterReplace
                            ),
                            "Замена завершена успешно!",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information
                            );
        }

        //Метод замены строки
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
