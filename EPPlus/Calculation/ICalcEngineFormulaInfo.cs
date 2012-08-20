using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Calculation
{
    public interface ICalcEngineFormulaInfo
    {
        Dictionary<string, string> GetFormulas();   //CellID?, Formula

        Dictionary<string, object> GetNameValues();
    }
}
