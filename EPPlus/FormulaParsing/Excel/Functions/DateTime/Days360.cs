using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Days360 : ExcelFunction
    {
        private enum Days360Calctype
        {
            European,
            Us
        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var numDate1 = ArgToDecimal(arguments, 0);
            var numDate2 = ArgToDecimal(arguments, 1);
            var dt1 = System.DateTime.FromOADate(numDate1);
            var dt2 = System.DateTime.FromOADate(numDate2);

            var calcType = Days360Calctype.Us;
            if (arguments.Count() > 2)
            {
                var european = ArgToBool(arguments, 2);
                if(european) calcType = Days360Calctype.European;
            }

            var startYear = dt1.Year;
            var startMonth = dt1.Month;
            var startDay = dt1.Day;
            var endYear = dt2.Year;
            var endMonth = dt2.Month;
            var endDay = dt2.Day;

            if (calcType == Days360Calctype.European)
            {
                if (startDay == 31) startDay = 30;
                if (endDay == 31) endDay = 30;
            }
            else
            {
                var calendar = new GregorianCalendar();
                var nDaysInFeb = calendar.IsLeapYear(dt1.Year) ? 29 : 28;
               
                 // If the investment is EOM and (Date1 is the last day of February) and (Date2 is the last day of February), then change D2 to 30.
                if (startMonth == 2 && startDay == nDaysInFeb && endMonth == 2 && endDay == nDaysInFeb)
                {
                    endDay = 30;
                }
                 // If the investment is EOM and (Date1 is the last day of February), then change D1 to 30.
                if (startMonth == 2 && startDay == nDaysInFeb)
                {
                    startDay = 30;
                }
                 // If D2 is 31 and D1 is 30 or 31, then change D2 to 30.
                if (endDay == 31 && (startDay == 30 || startDay == 31))
                {
                    endDay = 30;
                }
                 // If D1 is 31, then change D1 to 30.
                if (startDay == 31)
                {
                    startDay = 30;
                }
            }
            var result = (endYear*12*30 + endMonth*30 + endDay) - (startYear*12*30 + startMonth*30 + startDay);
            return CreateResult(result, DataType.Integer);
        }

        private int GetNumWholeMonths(System.DateTime dt1, System.DateTime dt2)
        {
            var startDate = new System.DateTime(dt1.Year, dt1.Month, 1).AddMonths(1);
            var endDate = new System.DateTime(dt2.Year, dt2.Month, 1);
            return ((endDate.Year - startDate.Year)*12) + (endDate.Month - startDate.Month);
        }
    }
}
