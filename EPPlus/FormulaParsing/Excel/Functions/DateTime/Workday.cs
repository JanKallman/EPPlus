using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            resultDate = AdjustResultWithHolidays(tmpDate, arguments);
            return CreateResult(resultDate.ToOADate(), DataType.Date);
        }

        private System.DateTime AdjustResultWithHolidays(System.DateTime resultDate,
                                                         IEnumerable<FunctionArgument> arguments)
        {
            if (arguments.Count() == 2) return resultDate;
            var holidays = arguments.ElementAt(2).Value as IEnumerable<FunctionArgument>;
            if (holidays != null)
            {
                foreach (var arg in holidays)
                {
                    if (ConvertUtil.IsNumeric(arg.Value))
                    {
                        var dateSerial = ConvertUtil.GetValueDouble(arg.Value);
                        var holidayDate = System.DateTime.FromOADate(dateSerial);
                        if (!IsHoliday(holidayDate))
                        {
                            resultDate = resultDate.AddDays(1);
                        }
                    }
                }
            }
            else
            {
                var range = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
                if (range != null)
                {
                    foreach (var cell in range)
                    {
                        if (ConvertUtil.IsNumeric(cell.Value))
                        {
                            var dateSerial = ConvertUtil.GetValueDouble(cell.Value);
                            var holidayDate = System.DateTime.FromOADate(dateSerial);
                            if (!IsHoliday(holidayDate))
                            {
                                resultDate = resultDate.AddDays(1);
                            }
                        }
                    }
                }
            }
            return resultDate;
        }

        private bool IsHoliday(System.DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        }
    }
}
