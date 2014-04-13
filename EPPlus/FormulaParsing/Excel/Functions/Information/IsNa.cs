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
            if (arguments == null || arguments.Count() == 0)
            {
                return CreateResult(false, DataType.Boolean);
            }

            var v = GetFirstValue(arguments);

            if (v is ExcelErrorValue && ((ExcelErrorValue)v).Type==eErrorType.NA)
            {
                return CreateResult(true, DataType.Boolean);
            }
            return CreateResult(false, DataType.Boolean);
        }
    }
}
