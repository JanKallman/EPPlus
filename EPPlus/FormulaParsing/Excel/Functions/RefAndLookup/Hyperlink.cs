using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class Hyperlink : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var link = ArgToString(arguments, 0);
            string friendlyName = null;
            if (arguments.Count() > 1)
            {
                friendlyName = ArgToString(arguments, 1);
            }
            return null;
        }
    }
}
