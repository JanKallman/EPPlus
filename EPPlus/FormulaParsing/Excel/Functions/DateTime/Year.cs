using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Year : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var dateObj = arguments.ElementAt(0).Value;
            System.DateTime date = System.DateTime.MinValue;
            if (dateObj is double)
            {
                date = System.DateTime.FromOADate((double)dateObj);
            }
            if (dateObj is string)
            {
                date = System.DateTime.Parse(dateObj.ToString());
            }
            return CreateResult(date.Year, DataType.Integer);
        }
    }
}
