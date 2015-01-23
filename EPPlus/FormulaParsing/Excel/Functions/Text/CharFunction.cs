using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class CharFunction : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToInt(arguments, 0);
            ThrowExcelErrorValueExceptionIf(() => number < 1 || number > 255, eErrorType.Value);
            return CreateResult(((char) number).ToString(), DataType.String);
        }
    }
}
