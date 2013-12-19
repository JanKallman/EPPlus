using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Exceptions
{
    public class ExcelErrorValueException : Exception
    {
        
        public ExcelErrorValueException(ExcelErrorValue error)
            : this(error.ToString(), error)
        {
            ErrorValue = error;
        }

        public ExcelErrorValueException(string message, ExcelErrorValue error)
            : base(message)
        {
            
        }

        public ExcelErrorValue ErrorValue { get; private set; }
    }
}
