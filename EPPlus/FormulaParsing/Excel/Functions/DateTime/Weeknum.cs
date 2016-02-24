using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Weeknum : ExcelFunction
    {
        public override ExpressionGraph.CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1, eErrorType.Value);
            var dateSerial = ArgToDecimal(arguments, 0);
            var date = System.DateTime.FromOADate(dateSerial);
            var startDay = DayOfWeek.Sunday;
            if (arguments.Count() > 1)
            {
                var argStartDay = ArgToInt(arguments, 1);
                switch (argStartDay)
                {
                    case 1:
                        startDay = DayOfWeek.Sunday;
                        break;
                    case 2:
                    case 11:
                        startDay = DayOfWeek.Monday;
                        break;
                    case 12:
                        startDay = DayOfWeek.Tuesday;
                        break;
                    case 13:
                        startDay = DayOfWeek.Wednesday;
                        break;
                    case 14:
                        startDay = DayOfWeek.Thursday;
                        break;
                    case 15:
                        startDay = DayOfWeek.Friday;
                        break;
                    case 16:
                        startDay = DayOfWeek.Saturday;
                        break;
                    default:
                        // Not supported 
                        ThrowExcelErrorValueException(eErrorType.Num);
                        break;
                }
            }
            if (DateTimeFormatInfo.CurrentInfo == null)
            {
                throw new InvalidOperationException(
                    "Could not execute Weeknum function because DateTimeFormatInfo.CurrentInfo was null");
            }
            var week = DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstDay,
                                                                             startDay);
            return CreateResult(week, DataType.Integer);
        }
        
        
    }
}
