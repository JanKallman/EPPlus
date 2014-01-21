using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    public class IsNa : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var isError = new IsError();
            var result = isError.Execute(arguments, context);
            if ((bool)result.Result)
            {
                if (arguments.ElementAt(0).Value.ToString() == ExcelErrorValue.Values.NA)
                {
                    return CreateResult(true, DataType.Boolean);
                }
            }
            return CreateResult(false, DataType.Boolean);
        }
    }
}
