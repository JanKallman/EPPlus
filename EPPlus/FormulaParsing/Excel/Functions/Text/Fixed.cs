using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Fixed : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var number = ArgToDecimal(arguments, 0);
            var nDecimals = 2;
            var noCommas = false;
            if (arguments.Count() > 1)
            {
                nDecimals = ArgToInt(arguments, 1);
            }
            if (arguments.Count() > 2)
            {
                noCommas = ArgToBool(arguments, 2);
            }
            var format = (noCommas ? "F" : "N") + nDecimals.ToString(CultureInfo.InvariantCulture);
            if (nDecimals < 0)
            {
                number = number - (number % (System.Math.Pow(10, nDecimals * -1)));
                number = System.Math.Floor(number);
                format = noCommas ? "F0" : "N0";
            }
            var retVal = number.ToString(format);
            return CreateResult(retVal, DataType.String);
        }
    }
}
