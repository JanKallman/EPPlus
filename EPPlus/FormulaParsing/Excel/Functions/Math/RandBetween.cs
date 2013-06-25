using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class RandBetween : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var low = ArgToDecimal(arguments, 0);
            var high = ArgToDecimal(arguments, 1);
            var rand = new Rand().Execute(new FunctionArgument[0], context).Result;
            var randPart = (CalulateDiff(high, low) * (double)rand) + 1;
            randPart = System.Math.Floor(randPart);
            return CreateResult(low + randPart, DataType.Integer);
        }

        private double CalulateDiff(double high, double low)
        {
            if (high > 0 && low < 0)
            {
                return high + low * - 1;
            }
            else if (high < 0 && low < 0)
            {
                return high * -1 - low * -1;
            }
            return high - low;
        }
    }
}
