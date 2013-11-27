using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Substitute : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var text = ArgToString(arguments, 0);
            var find = ArgToString(arguments, 1);
            var replaceWith = ArgToString(arguments, 2);
            var result = text.Replace(find, replaceWith);
            return CreateResult(result, DataType.String);
        }
    }
}
