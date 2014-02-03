using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    public class ExcelCell
    {
        public ExcelCell(object val, string formula, int colIndex, int rowIndex)
        {
            Value = val;
            Formula = formula;
            ColIndex = colIndex;
            RowIndex = rowIndex;
        }

        public int ColIndex { get; private set; }

        public int RowIndex { get; private set; }

        public object Value { get; private set; }

        public string Formula { get; private set; }
    }
}
