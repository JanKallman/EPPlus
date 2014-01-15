using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Exceptions
{
    public class ExcelErrorValueException : Exception
    {
        
        public ExcelErrorValueException(ExcelErrorValue error)
        {
            ErrorValue = error;
        }
        public ExcelErrorValue ErrorValue { get; private set; }
    }
}
