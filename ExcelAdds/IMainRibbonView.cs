using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelAdds
{
    public interface IMainRibbonView
    {
        string TargetString { get; }
        string Replacement { get; }
        string SelectedRange { set; }

        event Action ButtonReplaceClicked;
        event Action ButtonSetReplacementRangeClicked;
        event Action ButtonReplaceSubStringClicked;
    }
}
