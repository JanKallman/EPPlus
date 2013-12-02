using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Weekday : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var serialNumber = ArgToDecimal(arguments, 0);
            var returnType = ArgToInt(arguments, 1);
            return CreateResult(CalculateDayOfWeek(System.DateTime.FromOADate(serialNumber), returnType), DataType.Integer);
        }

        private static List<int> _oneBasedStartOnSunday = new List<int> { 1, 2, 3, 4, 5, 6, 7 };
        private static List<int> _oneBasedStartOnMonday = new List<int> { 7, 1, 2, 3, 4, 5, 6 };
        private static List<int> _zeroBasedStartOnSunday = new List<int> { 6, 0, 1, 2, 3, 4, 5 };

        private int CalculateDayOfWeek(System.DateTime dateTime, int returnType)
        {
            var dayIx = (int)dateTime.DayOfWeek;
            switch (returnType)
            {
                case 1:
                    return _oneBasedStartOnSunday[dayIx];
                case 2:
                    return _oneBasedStartOnMonday[dayIx];
                case 3:
                    return _zeroBasedStartOnSunday[dayIx];
                default:
                    throw new ArgumentException("invalid return type (should be 1-3) supplied to Weekday function: " + returnType);
            }
        }
    }
}
