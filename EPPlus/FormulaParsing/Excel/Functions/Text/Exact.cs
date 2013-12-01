using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Exact : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var val1 = arguments.ElementAt(0).Value;
            var val2 = arguments.ElementAt(1).Value;

            if (val1 == null && val2 == null)
            {
                return CreateResult(true, DataType.Boolean);
            }
            else if ((val1 == null && val2 != null) || (val1 != null && val2 == null))
            {
                return CreateResult(false, DataType.Boolean);
            }

            var result = string.Compare(val1.ToString(), val2.ToString(), StringComparison.InvariantCulture);
            return CreateResult(result == 0, DataType.Boolean);
        }
    }
}
