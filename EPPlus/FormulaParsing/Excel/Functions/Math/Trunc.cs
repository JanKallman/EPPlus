using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Trunc : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToDecimal(arguments, 0);
            if (arguments.Count() == 1)
            {
                return CreateResult(System.Math.Truncate(number), DataType.Decimal);
            }
            var nDigits = ArgToInt(arguments, 1);
            var func = context.Configuration.FunctionRepository.GetFunction("rounddown");
            return func.Execute(arguments, context);
        }
    }
}
