using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Atan2 : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var arg1 = ArgToDecimal(arguments, 0);
            var arg2 = ArgToDecimal(arguments, 1);
            // Had to switch order of the arguments to get the same result as in excel /MA
            return CreateResult(System.Math.Atan2(arg2, arg1), DataType.Decimal);
        }
    }
}
