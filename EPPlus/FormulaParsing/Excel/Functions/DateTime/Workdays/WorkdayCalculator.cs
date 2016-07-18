using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays
{
    public class WorkdayCalculator
    {
        private readonly HolidayWeekdays _holidayWeekdays;

        public WorkdayCalculator()
            : this(new HolidayWeekdays())
        {}

        public WorkdayCalculator(HolidayWeekdays holidayWeekdays)
        {
            _holidayWeekdays = holidayWeekdays;
        }

        public System.DateTime CalculateWorkday(System.DateTime startDate, int nWorkDays)
        {
            var direction = nWorkDays > 0 ? 1 : -1;
            nWorkDays *= direction;
            var workdaysCounted = 0;
            var tmpDate = startDate;
            // first move forward to the first monday
            while (tmpDate.DayOfWeek != DayOfWeek.Monday && (nWorkDays - workdaysCounted) > 0)
            {
                if (!_holidayWeekdays.IsHolidayWeekday(tmpDate)) workdaysCounted++;
                tmpDate = tmpDate.AddDays(direction);
            }
            // then calculate whole weeks
            var nWholeWeeks = (nWorkDays - workdaysCounted) / _holidayWeekdays.NumberOfWorkdaysPerWeek;
            tmpDate = tmpDate.AddDays(nWholeWeeks * 7 * direction);
            workdaysCounted += nWholeWeeks * _holidayWeekdays.NumberOfWorkdaysPerWeek;

            // calculate the rest
            while (workdaysCounted < nWorkDays)
            {
                tmpDate = tmpDate.AddDays(direction);
                if (!_holidayWeekdays.IsHolidayWeekday(tmpDate)) workdaysCounted++;
            }
            return tmpDate;
        }

        public System.DateTime AdjustResultWithHolidays(System.DateTime startDate, System.DateTime resultDate,
                                                         IEnumerable<FunctionArgument> arguments, bool forward)
        {

            #region old code
            /*
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
                        if (!_holidayWeekdays.IsHolidayWeekday(holidayDate))
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
                            if (!_holidayWeekdays.IsHolidayWeekday(holidayDate))
                            {
                                resultDate = resultDate.AddDays(1);
                            }
                        }
                    }
                }
            }
            */
            #endregion

            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            if (functionArguments.Count() > 2)
            {
                var additionalDays = new AdditionalHolidayDays(functionArguments.ElementAt(2));
                foreach (var date in additionalDays.AdditionalDates)
                {
                    if (forward && (date < startDate || date > resultDate)) continue;
                    if (!forward && (date > startDate || date < resultDate)) continue;
                    if (_holidayWeekdays.IsHolidayWeekday(date)) continue;
                    var tmpDate = _holidayWeekdays.GetNextWorkday(resultDate, forward);
                    while (additionalDays.AdditionalDates.Contains(tmpDate))
                    {
                        tmpDate = _holidayWeekdays.GetNextWorkday(tmpDate, forward);
                    }
                    resultDate = tmpDate;
                }
            }

            return resultDate;
        }
    }
}
