using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    /// <summary>
    /// Simple implementation of DateValue function, just using .NET built-in
    /// function System.DateTime.TryParse, based on current culture
    /// </summary>
    public class DateValue : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var dateString = ArgToString(arguments, 0);
            return Execute(dateString);
        }

        internal CompileResult Execute(string dateString)
        {
            System.DateTime result;
            System.DateTime.TryParse(dateString, out result);
            return result != System.DateTime.MinValue ?
                CreateResult(result.ToOADate(), DataType.Date) :
                CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
        }
    }
}
