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
            var date1num = ArgToDecimal(functionArguments, 0);
            var date2num = ArgToDecimal(functionArguments, 1);
            var date1 = System.DateTime.FromOADate(date1num);
            var date2 = System.DateTime.FromOADate(date2num);
            var basis = 0;
            if (functionArguments.Count() > 2)
            {
                basis = ArgToInt(functionArguments, 2);
                ThrowExcelErrorValueExceptionIf(() => basis < 0 || basis > 4, eErrorType.Num);
            }
            var func = context.Configuration.FunctionRepository.GetFunction("days360");
            var calendar = new GregorianCalendar();
            double? result;
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
                case 2:
                case 3:
                    throw new NotImplementedException();
                case 4:
                    var args = functionArguments.ToList();
                    args.Add(new FunctionArgument(true));
                    result = func.Execute(args, context).ResultNumeric/360d;
                    return CreateResult(result.Value, DataType.Decimal);
                default:
                    return null;
            }
        }
    }
}
