using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Log : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToDecimal(arguments, 0);
            if (arguments.Count() == 1)
            {
                return CreateResult(System.Math.Log(number, 10d), DataType.Decimal);
            }
            var newBase = ArgToDecimal(arguments, 1);
            return CreateResult(System.Math.Log(number, newBase), DataType.Decimal);
        }
    }
}
