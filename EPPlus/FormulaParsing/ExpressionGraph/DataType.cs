using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public enum DataType
    {
        Integer,
        Decimal,
        String,
        Boolean,
        Date,
        Time,
        Enumerable,
        LookupArray,
        ExcelAddress,
        Empty
    }
}
