using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Value : ExcelFunction
    {
        private readonly string _groupSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
        private readonly string _decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        private readonly string _timeSeparator = CultureInfo.CurrentCulture.DateTimeFormat.TimeSeparator;
        private readonly string _shortTimePattern = CultureInfo.CurrentCulture.DateTimeFormat.ShortTimePattern;
        private readonly DateValue _dateValueFunc = new DateValue();
        private readonly TimeValue _timeValueFunc = new TimeValue();

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var val = ArgToString(arguments, 0).TrimEnd(' ');
            double result = 0d;
            if (Regex.IsMatch(val, $"^[\\d]*({Regex.Escape(_groupSeparator)}?[\\d]*)?({Regex.Escape(_decimalSeparator)}[\\d]*)?[ ?% ?]?$"))
            {
                if (val.EndsWith("%"))
                {
                    val = val.TrimEnd('%');
                    result = double.Parse(val) / 100;
                }
                else
                {
                    result = double.Parse(val);
                }
                return CreateResult(result, DataType.Decimal);
            }
            if (double.TryParse(val, NumberStyles.Float, CultureInfo.CurrentCulture, out result))
            {
                return CreateResult(result, DataType.Decimal);
            }
            var timeSeparator = Regex.Escape(_timeSeparator);
            if (Regex.IsMatch(val, @"^[\d]{1,2}" + timeSeparator + @"[\d]{2}(" + timeSeparator + @"[\d]{2})?$"))
            {
                var timeResult = _timeValueFunc.Execute(val);
                if (timeResult.DataType == DataType.Date)
                {
                    return timeResult;
                }
            }
            var dateResult = _dateValueFunc.Execute(val);
            if (dateResult.DataType == DataType.Date)
            {
                return dateResult;
            }
            return CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
        }
    }
}
