using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Hyperlink : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            if (arguments.Count() > 1)
            {
                return CreateResult(ArgToString(arguments, 1), DataType.String);
            }
            return CreateResult(ArgToString(arguments, 0), DataType.String);
        }
    }
}
