using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays
{
    public class HolidayWeekdaysFactory
    {
        private readonly DayOfWeek[] _dayOfWeekArray = new DayOfWeek[]
        {
            DayOfWeek.Monday, 
            DayOfWeek.Tuesday, 
            DayOfWeek.Wednesday, 
            DayOfWeek.Thursday,
            DayOfWeek.Friday, 
            DayOfWeek.Saturday,
            DayOfWeek.Sunday
        };

        public HolidayWeekdays Create(string weekdays)
        {
            if(string.IsNullOrEmpty(weekdays) || weekdays.Length != 7)
                throw new ArgumentException("Illegal weekday string", nameof(Weekday));

            var retVal = new List<DayOfWeek>();
            var arr = weekdays.ToCharArray();
            for(var i = 0; i < arr.Length;i++)
            {
                var ch = arr[i];
                if (ch == '1')
                {
                    retVal.Add(_dayOfWeekArray[i]);
                }
            }
            return new HolidayWeekdays(retVal.ToArray());
        }

        public HolidayWeekdays Create(int code)
        {
            switch (code)
            {
                case 1:
                    return new HolidayWeekdays(DayOfWeek.Saturday, DayOfWeek.Sunday);
                case 2:
                    return new HolidayWeekdays(DayOfWeek.Sunday, DayOfWeek.Monday);
                case 3:
                    return new HolidayWeekdays(DayOfWeek.Monday, DayOfWeek.Tuesday);
                case 4:
                    return new HolidayWeekdays(DayOfWeek.Tuesday, DayOfWeek.Wednesday);
                case 5:
                    return new HolidayWeekdays(DayOfWeek.Wednesday, DayOfWeek.Thursday);
                case 6:
                    return new HolidayWeekdays(DayOfWeek.Thursday, DayOfWeek.Friday);
                case 7:
                    return new HolidayWeekdays(DayOfWeek.Friday, DayOfWeek.Saturday);
                case 11:
                    return new HolidayWeekdays(DayOfWeek.Sunday);
                case 12:
                    return new HolidayWeekdays(DayOfWeek.Monday);
                case 13:
                    return new HolidayWeekdays(DayOfWeek.Tuesday);
                case 14:
                    return new HolidayWeekdays(DayOfWeek.Wednesday);
                case 15:
                    return new HolidayWeekdays(DayOfWeek.Thursday);
                case 16:
                    return new HolidayWeekdays(DayOfWeek.Friday);
                case 17:
                    return new HolidayWeekdays(DayOfWeek.Saturday);
                default:
                    throw new ArgumentException("Invalid code supplied to HolidayWeekdaysFactory: " + code);
            }
        }
    }
}
