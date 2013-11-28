using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class Choose : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var index = ArgToInt(arguments, 0);
            var items = new List<string>();
            for (int x = 0; x < arguments.Count(); x++)
            {
                items.Add(arguments.ElementAt(x).Value.ToString());
            }
            return CreateResult(items[index], DataType.String);
        }
    }
}
