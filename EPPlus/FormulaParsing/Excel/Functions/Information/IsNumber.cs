using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    public class IsNumber : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = arguments.ElementAt(0).Value;
            if (arg is System.DateTime || arg is double || arg is int || arg is decimal || arg is short || arg is long)
            {
                return CreateResult(true, DataType.Boolean);
            }
            return CreateResult(false, DataType.Boolean);
        }
    }
}
