using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public abstract class ErrorHandlingFunction : ExcelFunction
    {
        public override bool IsErrorHandlingFunction
        {
            get
            {
                return true;
            }
        }

        public abstract CompileResult HandleError(string errorCode);
    }
}
