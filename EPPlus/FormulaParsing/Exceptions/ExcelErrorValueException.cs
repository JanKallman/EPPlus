using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Exceptions
{
    /// <summary>
    /// This Exception represents an Excel error. When this exception is thrown
    /// from an Excel function, the ErrorValue code will be set as the value of the
    /// parsed cell.
    /// </summary>
    /// <seealso cref="ExcelErrorValue"/>
    public class ExcelErrorValueException : Exception
    {
        
        public ExcelErrorValueException(ExcelErrorValue error)
            : this(error.ToString(), error)
        {
            
        }

        public ExcelErrorValueException(string message, ExcelErrorValue error)
            : base(message)
        {
            ErrorValue = error;
        }

        public ExcelErrorValueException(eErrorType errorType)
            : this(ExcelErrorValue.Create(errorType))
        {
            
        }

        /// <summary>
        /// The error value
        /// </summary>
        public ExcelErrorValue ErrorValue { get; private set; }
    }
}
