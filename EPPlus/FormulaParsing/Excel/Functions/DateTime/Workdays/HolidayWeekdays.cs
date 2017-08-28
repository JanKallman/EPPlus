using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays
{
    public class HolidayWeekdays
    {
        private readonly List<DayOfWeek> _holidayDays = new List<DayOfWeek>();

        public HolidayWeekdays()
            :this(DayOfWeek.Saturday, DayOfWeek.Sunday)
        {
            
        }

        public int NumberOfWorkdaysPerWeek => 7 - _holidayDays.Count;

        public HolidayWeekdays(params DayOfWeek[] holidayDays)
        {
            foreach (var dayOfWeek in holidayDays)
            {
                _holidayDays.Add(dayOfWeek);
            }
        }

        public bool IsHolidayWeekday(System.DateTime dateTime)
        {
            return _holidayDays.Contains(dateTime.DayOfWeek);
        }

        public System.DateTime AdjustResultWithHolidays(System.DateTime resultDate,
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
                        if (!IsHolidayWeekday(holidayDate))
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
                            if (!IsHolidayWeekday(holidayDate))
                            {
                                resultDate = resultDate.AddDays(1);
                            }
                        }
                    }
                }
            }
            return resultDate;
        }

        public System.DateTime GetNextWorkday(System.DateTime date, WorkdayCalculationDirection direction = WorkdayCalculationDirection.Forward)
        {
            var changeParam = (int)direction;
            var tmpDate = date.AddDays(changeParam);
            while (IsHolidayWeekday(tmpDate))
            {
                tmpDate = tmpDate.AddDays(changeParam);
            }
            return tmpDate;
        }
    }
}
