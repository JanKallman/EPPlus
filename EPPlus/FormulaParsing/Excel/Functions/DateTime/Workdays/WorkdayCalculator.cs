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

        public WorkdayCalculatorResult CalculateNumberOfWorkdays(System.DateTime startDate, System.DateTime endDate)
        {
            var calcDirection = startDate < endDate
                ? WorkdayCalculationDirection.Forward
                : WorkdayCalculationDirection.Backward;
            System.DateTime calcStartDate;
            System.DateTime calcEndDate;
            if (calcDirection == WorkdayCalculationDirection.Forward)
            {
                calcStartDate = startDate.Date;
                calcEndDate = endDate.Date;
            }
            else
            {
                calcStartDate = endDate.Date;
                calcEndDate = startDate.Date;
            }
            var nWholeWeeks = (int)calcEndDate.Subtract(calcStartDate).TotalDays/7;
            var workdaysCounted = nWholeWeeks*_holidayWeekdays.NumberOfWorkdaysPerWeek;
            if (!_holidayWeekdays.IsHolidayWeekday(calcStartDate))
            {
                workdaysCounted++;
            }
            var tmpDate = calcStartDate.AddDays(nWholeWeeks*7);
            while (tmpDate < calcEndDate)
            {
                tmpDate = tmpDate.AddDays(1);
                if (!_holidayWeekdays.IsHolidayWeekday(tmpDate))
                {
                    workdaysCounted++;
                }
            }
            return new WorkdayCalculatorResult(workdaysCounted, startDate, endDate, calcDirection);
        }

        public WorkdayCalculatorResult CalculateWorkday(System.DateTime startDate, int nWorkDays)
        {
            var calcDirection = nWorkDays > 0 ? WorkdayCalculationDirection.Forward : WorkdayCalculationDirection.Backward;
            var direction = (int) calcDirection;
            nWorkDays *= direction;
            var workdaysCounted = 0;
            var tmpDate = startDate;
            
            // calculate whole weeks
            var nWholeWeeks = nWorkDays / _holidayWeekdays.NumberOfWorkdaysPerWeek;
            tmpDate = tmpDate.AddDays(nWholeWeeks * 7 * direction);
            workdaysCounted += nWholeWeeks * _holidayWeekdays.NumberOfWorkdaysPerWeek;

            // calculate the rest
            while (workdaysCounted < nWorkDays)
            {
                tmpDate = tmpDate.AddDays(direction);
                if (!_holidayWeekdays.IsHolidayWeekday(tmpDate)) workdaysCounted++;
            }
            return new WorkdayCalculatorResult(workdaysCounted, startDate, tmpDate, calcDirection);
        }

        public WorkdayCalculatorResult ReduceWorkdaysWithHolidays(WorkdayCalculatorResult calculatedResult,
            FunctionArgument holidayArgument)
        {
            var startDate = calculatedResult.StartDate;
            var endDate = calculatedResult.EndDate;
            var additionalDays = new AdditionalHolidayDays(holidayArgument);
            System.DateTime calcStartDate;
            System.DateTime calcEndDate;
            if (startDate < endDate)
            {
                calcStartDate = startDate;
                calcEndDate = endDate;
            }
            else
            {
                calcStartDate = endDate;
                calcEndDate = startDate;
            }
            var nAdditionalHolidayDays = additionalDays.AdditionalDates.Count(x => x >= calcStartDate && x <= calcEndDate && !_holidayWeekdays.IsHolidayWeekday(x));
            return new WorkdayCalculatorResult(calculatedResult.NumberOfWorkdays - nAdditionalHolidayDays, startDate, endDate, calculatedResult.Direction);
        } 

        public WorkdayCalculatorResult AdjustResultWithHolidays(WorkdayCalculatorResult calculatedResult,
                                                         FunctionArgument holidayArgument)
        {
            var startDate = calculatedResult.StartDate;
            var endDate = calculatedResult.EndDate;
            var direction = calculatedResult.Direction;
            var workdaysCounted = calculatedResult.NumberOfWorkdays;
            var additionalDays = new AdditionalHolidayDays(holidayArgument);
            foreach (var date in additionalDays.AdditionalDates)
            {
                if (direction == WorkdayCalculationDirection.Forward && (date < startDate || date > endDate)) continue;
                if (direction == WorkdayCalculationDirection.Backward && (date > startDate || date < endDate)) continue;
                if (_holidayWeekdays.IsHolidayWeekday(date)) continue;
                var tmpDate = _holidayWeekdays.GetNextWorkday(endDate, direction);
                while (additionalDays.AdditionalDates.Contains(tmpDate))
                {
                    tmpDate = _holidayWeekdays.GetNextWorkday(tmpDate, direction);
                }
                workdaysCounted++;
                endDate = tmpDate;
            }

            return new WorkdayCalculatorResult(workdaysCounted, calculatedResult.StartDate, endDate, direction);
        }
    }
}
