using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    public class ParsedValue
    {
        public ParsedValue(object val, int rowIndex, int colIndex)
        {
            Value = val;
            RowIndex = rowIndex;
            ColIndex = colIndex;
        }

        public object Value
        {
            get;
            private set;
        }

        public int RowIndex
        {
            get;
            private set;
        }

        public int ColIndex
        {
            get;
            private set;
        }
    }
}
