using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAdds
{
    public partial class MainRibbon : IMainRibbonView
    {
        public event System.Action ButtonReplaceClicked;
        public event System.Action ButtonSetReplacementRangeClicked;
        public event System.Action ButtonReplaceSubStringClicked;

        public string TargetString { get { return txbTarget.Text; } }
        public string Replacement { get { return txbReplacement.Text; } }
        public string SelectedRange { set { txbSelectedRange.Text = value; } }

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnReplace_Click(object sender, RibbonControlEventArgs e)
        {
            if (ButtonReplaceClicked != null) ButtonReplaceClicked();
        }

        private void btnSetReplacementRange_Click(object sender, RibbonControlEventArgs e)
        {
            if (ButtonSetReplacementRangeClicked != null) ButtonSetReplacementRangeClicked();
        }

        private void btnReplaceSubString_Click(object sender, RibbonControlEventArgs e)
        {
            if (ButtonReplaceSubStringClicked != null) ButtonReplaceSubStringClicked();
        }


    }
}
