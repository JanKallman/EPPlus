using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Concatenate : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments == null)
            {
                return CreateResult(string.Empty, DataType.String);
            }
            var sb = new StringBuilder();
            foreach (var arg in arguments)
            {
                sb.Append(arg.Value.ToString());
            }
            return CreateResult(sb.ToString(), DataType.String);
        }
    }
}
