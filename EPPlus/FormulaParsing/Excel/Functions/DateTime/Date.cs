using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Date : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var year = ArgToInt(arguments, 0);
            var month = ArgToInt(arguments, 1);
            var day = ArgToInt(arguments, 2);
            var date = new System.DateTime(year, 1, 1);
            month -= 1;
            date = date.AddMonths(month);
            date = date.AddDays((double)(day - 1));
            return CreateResult(date.ToOADate(), DataType.Date);
        }
    }
}
