using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel
{
    [Flags]
    public enum ExcelCellState
    {
        HiddenCell = 1,
        ContainsError = 2
    }
}
