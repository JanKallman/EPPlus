using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Mid : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var text = ArgToString(arguments, 0);
            var startIx = ArgToInt(arguments, 1);
            var length = ArgToInt(arguments, 2);
            var result = text.Substring(startIx, length);
            return CreateResult(result, DataType.String);
        }
    }
}
