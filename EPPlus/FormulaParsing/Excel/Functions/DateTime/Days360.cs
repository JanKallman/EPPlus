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

            int nDaysInStartMonth = 0, nDaysInEndMonth = 0;
            nDaysInStartMonth = (dt1.Day == 31) ? 1 : 30 - dt1.Day;
            if (calcType == Days360Calctype.European)
            { 
                nDaysInEndMonth = dt2.Day == 31 ? 30 : dt2.Day;
            }
            else
            {
                nDaysInEndMonth = dt2.Day;
                if (dt1.Day == 31 || dt1.Day == 30)
                {
                    nDaysInEndMonth++;
                }
                if (dt1.Month == 2 && dt1.Day > 27 && dt2.Month == 2 && dt2.Day > 27)
                {
                    nDaysInEndMonth -= (30 - dt2.Day);
                }
                else if (dt1.Month == 2 && dt1.Day > 27)
                {
                    var calendar = new GregorianCalendar();
                    var nDaysInFeb = calendar.IsLeapYear(dt1.Year) ? 29 : 28;
                    if (dt1.Day == nDaysInFeb)
                    {
                        nDaysInStartMonth -= (31 - dt1.Day);
                    }
                }
            }

            var result = nDaysInStartMonth + GetNumWholeMonths(dt1, dt2) * 30 + nDaysInEndMonth;

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
