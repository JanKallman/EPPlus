using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    public class Yearfrac : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 2);
            var date1Num = ArgToDecimal(functionArguments, 0);
            var date2Num = ArgToDecimal(functionArguments, 1);
            var date1 = System.DateTime.FromOADate(date1Num);
            var date2 = System.DateTime.FromOADate(date2Num);
            var basis = 0;
            if (functionArguments.Count() > 2)
            {
                basis = ArgToInt(functionArguments, 2);
                ThrowExcelErrorValueExceptionIf(() => basis < 0 || basis > 4, eErrorType.Num);
            }
            var func = context.Configuration.FunctionRepository.GetFunction("days360");
            var calendar = new GregorianCalendar();
            switch (basis)
            {
                case 0:
                    var d360Result = func.Execute(functionArguments, context).ResultNumeric;
                    // reproducing excels behaviour
                    if (date1.Month == 2)
                    {
                        var daysInFeb = calendar.IsLeapYear(date1.Year) ? 29 : 28;
                        if (date1.Day == daysInFeb) d360Result++;  
                    }
                    return CreateResult(d360Result / 360d, DataType.Decimal);
                case 1:
                    return CreateResult((date2 - date1).TotalDays/CalculateAcutalYear(date1, date2), DataType.Decimal);
                case 2:
                    return CreateResult((date2 - date1).TotalDays / 360d, DataType.Decimal);
                case 3:
                    return CreateResult((date2 - date1).TotalDays / 365d, DataType.Decimal);
                case 4:
                    var args = functionArguments.ToList();
                    args.Add(new FunctionArgument(true));
                    double? result = func.Execute(args, context).ResultNumeric/360d;
                    return CreateResult(result.Value, DataType.Decimal);
                default:
                    return null;
            }
        }

        private double CalculateAcutalYear(System.DateTime dt1, System.DateTime dt2)
        {
            var calendar = new GregorianCalendar();
            var perYear = 0d;
            var nYears = dt2.Year - dt1.Year + 1;
            for (var y = dt1.Year; y <= dt2.Year; ++y)
            {
                perYear += calendar.IsLeapYear(y) ? 366 : 365;
            }
            if (new System.DateTime(dt1.Year + 1, dt1.Month, dt1.Day) >= dt2)
            {
                nYears = 1;
                perYear = 365;
                if (calendar.IsLeapYear(dt1.Year) && dt1.Month <= 2)
                    perYear = 366;
                else if (calendar.IsLeapYear(dt2.Year) && dt2.Month > 2)
                    perYear = 366;
                else if (dt2.Month == 2 && dt2.Day == 29)
                    perYear = 366;
            }
            return perYear/(double) nYears;  
        }
    }
}
