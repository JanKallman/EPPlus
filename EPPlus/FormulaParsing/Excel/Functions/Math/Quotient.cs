using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Quotient : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var num = ArgToDecimal(arguments, 0);
            var denom = ArgToDecimal(arguments, 1);
            ThrowExcelErrorValueExceptionIf(() => (int)denom == 0, eErrorType.Div0);
            var result = (int)(num/denom);
            return CreateResult(result, DataType.Integer);
        }
    }
}
