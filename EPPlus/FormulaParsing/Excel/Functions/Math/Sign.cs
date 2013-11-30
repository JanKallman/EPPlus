using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Sign : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var result = 0d;
            var val = ArgToDecimal(arguments, 0);
            if (val < 0)
            {
                result = -1;
            }
            else if (val > 0)
            {
                result = 1;
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}
