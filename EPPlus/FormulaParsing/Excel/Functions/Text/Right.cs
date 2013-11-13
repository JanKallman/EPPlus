using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Right : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var str = ArgToString(arguments, 0);
            var length = ArgToInt(arguments, 1);
            var startIx = str.Length - length;
            return CreateResult(str.Substring(startIx, str.Length - startIx), DataType.String);
        }
    }
}
