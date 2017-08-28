using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Edate : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2, eErrorType.Value);
            var dateSerial = ArgToDecimal(arguments, 0);
            var date = System.DateTime.FromOADate(dateSerial);
            var nMonthsToAdd = ArgToInt(arguments, 1);
            var resultDate = date.AddMonths(nMonthsToAdd);
            return CreateResult(resultDate.ToOADate(), DataType.Date);
        }
    }
}
