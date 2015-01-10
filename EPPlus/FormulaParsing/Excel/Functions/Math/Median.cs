using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Median : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var nums = ArgsToDoubleEnumerable(arguments, context);
            var arr = nums.ToArray();
            Array.Sort(arr);
            ThrowExcelErrorValueExceptionIf(() => arr.Length == 0, eErrorType.Num);
            double result;
            if (arr.Length % 2 == 1)
            {
                result = arr[arr.Length / 2];
            }
            else
            {
                var startIndex = arr.Length/2 - 1;
                result = (arr[startIndex] + arr[startIndex + 1])/2d;
            }
            return CreateResult(result, DataType.Decimal);
        }
    }
}
