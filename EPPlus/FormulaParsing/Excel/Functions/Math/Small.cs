using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Small : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var args = arguments.ElementAt(0);
            var index = ArgToInt(arguments, 1) - 1;
            var values = ArgsToDoubleEnumerable(new List<FunctionArgument> { args }, context);
            ThrowExcelErrorValueExceptionIf(() => index < 0 || index >= values.Count(), eErrorType.Num);
            var result = values.OrderBy(x => x).ElementAt(index);
            return CreateResult(result, DataType.Decimal);
        }
    }
}
