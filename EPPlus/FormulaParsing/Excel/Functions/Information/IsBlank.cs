using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    public class IsBlank : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments == null || arguments.Count() == 0)
            {
                return CreateResult(true, DataType.Boolean);
            }
            var result = true;
            foreach (var arg in arguments)
            {
                if (arg.Value != null && (arg.Value.ToString() != string.Empty))
                {
                    result = false;
                    break;
                }
            }
            return CreateResult(result, DataType.Boolean);
        }
    }
}
