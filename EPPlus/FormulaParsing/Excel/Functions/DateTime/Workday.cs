using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

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
            var workdaysCounted = 0;
            var tmpDate = startDate;
            // first move forward to the first monday
            while (tmpDate.DayOfWeek != DayOfWeek.Monday && (nWorkDays - workdaysCounted) > 0)
            {
                if (!IsHoliday(tmpDate)) workdaysCounted++;
                tmpDate = tmpDate.AddDays(1);
            }
            // then calculate whole weeks
            var nWholeWeeks = (nWorkDays - workdaysCounted) / 5;
            tmpDate = tmpDate.AddDays(nWholeWeeks * 7);
            workdaysCounted += nWholeWeeks * 5;

            // calculate the rest
            while (workdaysCounted < nWorkDays)
            {
                tmpDate = tmpDate.AddDays(1);
                if (!IsHoliday(tmpDate)) workdaysCounted++;
            }
            resultDate = tmpDate;
            return CreateResult(resultDate.ToOADate(), DataType.Date);
        }

        private bool IsHoliday(System.DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }
    }
}
