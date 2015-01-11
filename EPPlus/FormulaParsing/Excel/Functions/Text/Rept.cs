using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Rept : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var str = ArgToString(arguments, 0);
            var n = ArgToInt(arguments, 1);
            var sb = new StringBuilder();
            for (var x = 0; x < n; x++)
            {
                sb.Append(str);
            }
            return CreateResult(sb.ToString(), DataType.String);
        }
    }
}
