using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Floor : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var number = ArgToDecimal(arguments, 0);
            var significance = ArgToDecimal(arguments, 1);
            ValidateNumberAndSign(number, significance);
            if (significance < 1 && significance > 0)
            {
                var floor = System.Math.Floor(number);
                var rest = number - floor;
                var nSign = (int)(rest / significance);
                return CreateResult(floor + (nSign * significance), DataType.Decimal);
            }
            else if (significance == 1)
            {
                return CreateResult(System.Math.Floor(number), DataType.Decimal);
            }
            else
            {
                double result;
                if (number > 1)
                {
                    result = number - (number % significance) + significance;
                }
                else
                {
                    result = number - (number % significance);
                }
                return CreateResult(result, DataType.Decimal);
            }
        }

        private void ValidateNumberAndSign(double number, double sign)
        {
            if (number > 0d && sign < 0)
            {
                var values = string.Format("num: {0}, sign: {1}", number, sign);
                throw new InvalidOperationException("Floor cannot handle a negative significance when the number is positive" + values);
            }
        }
    }
}
