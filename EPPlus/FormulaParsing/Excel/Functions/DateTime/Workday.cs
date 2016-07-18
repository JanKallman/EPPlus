using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Workday : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var startDate = System.DateTime.FromOADate(ArgToInt(arguments, 0));
            var nWorkDays = ArgToInt(arguments, 1);
            var resultDate = System.DateTime.MinValue;
            
            var calculator = new WorkdayCalculator();
            var tmpDate = calculator.CalculateWorkday(startDate, nWorkDays);
            resultDate = calculator.AdjustResultWithHolidays(startDate, tmpDate, arguments, nWorkDays > -1);
            return CreateResult(resultDate.ToOADate(), DataType.Date);
        }   
    }
}
