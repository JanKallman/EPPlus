using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class TimeStringParser
    {
        private const string RegEx24 = @"^[0-9]{1,2}(\:[0-9]{1,2}){0,2}$";
        private const string RegEx12 = @"^[0-9]{1,2}(\:[0-9]{1,2}){0,2}( PM| AM)$";

        private double GetSerialNumber(int hour, int minute, int second)
        {
            var secondsInADay = 24d * 60d * 60d;
            return ((double)hour * 60 * 60 + (double)minute * 60 + (double)second) / secondsInADay;
        }

        private void ValidateValues(int hour, int minute, int second)
        {
            if (second < 0 || second > 59)
            {
                throw new FormatException("Illegal value for second: " + second);
            }
            if (minute < 0 || minute > 59)
            {
                throw new FormatException("Illegal value for minute: " + minute);
            }
        }

        public virtual double Parse(string input)
        {
            return InternalParse(input);
        }

        public virtual bool CanParse(string input)
        {
            System.DateTime dt;
            return Regex.IsMatch(input, RegEx24) || Regex.IsMatch(input, RegEx12) || System.DateTime.TryParse(input, out dt);
        }

        private double InternalParse(string input)
        {
            if (Regex.IsMatch(input, RegEx24))
            {
                return Parse24HourTimeString(input);
            }
            if (Regex.IsMatch(input, RegEx12))
            {
                return Parse12HourTimeString(input);
            }
            System.DateTime dateTime;
            if (System.DateTime.TryParse(input, out dateTime))
            {
                return GetSerialNumber(dateTime.Hour, dateTime.Minute, dateTime.Second);
            }
            return -1;
        }

        private double Parse12HourTimeString(string input)
        {
            string dayPart = string.Empty;
            dayPart = input.Substring(input.Length - 2, 2);
            int hour;
            int minute;
            int second;
            GetValuesFromString(input, out hour, out minute, out second);
            if (dayPart == "PM") hour += 12;
            ValidateValues(hour, minute, second);
            return GetSerialNumber(hour, minute, second);
        }

        private double Parse24HourTimeString(string input)
        {
            int hour;
            int minute;
            int second;
            GetValuesFromString(input, out hour, out minute, out second);
            ValidateValues(hour, minute, second);
            return GetSerialNumber(hour, minute, second);
        }

        private static void GetValuesFromString(string input, out int hour, out int minute, out int second)
        {
            hour = 0;
            minute = 0;
            second = 0;

            var items = input.Split(':');
            hour = int.Parse(items[0]);
            if (items.Length > 1)
            {
                minute = int.Parse(items[1]);
            }
            if (items.Length > 2)
            {
                var val = items[2];
                val = Regex.Replace(val, "[^0-9]+$", string.Empty);
                second = int.Parse(val);
            }
        }
    }
}
