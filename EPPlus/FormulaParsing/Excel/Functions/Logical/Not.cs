using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    public class Not : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = arguments.ElementAt(0).Value;
            bool result = false;
            if (arg is bool)
            {
                result = !((bool)arg);
            }
            else if (arg is int)
            {
                result = ((int)arg) == 0;
            }
            return new CompileResult(result, DataType.Boolean);
        }
    }
}
