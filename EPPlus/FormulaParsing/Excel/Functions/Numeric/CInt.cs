using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric
{
    public class CInt : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var obj = arguments.ElementAt(0).Value;
            return CreateResult(ToInteger(obj), DataType.Integer);
        }

        private object ToInteger(object obj)
        {
            var type = obj.GetType();
            if (type == typeof(double))
            {
                return (int)System.Math.Floor((double)obj);
            }
            if (type == typeof(decimal))
            {
                return (int)System.Math.Floor((decimal)obj);
            }
            double result;
            if(double.TryParse(HandleDecimalSeparator(obj), out result))
            {
                return (int)System.Math.Floor(result);
            }
            throw new ArgumentException("Could not cast supplied argument to integer");
        }

        private string HandleDecimalSeparator(object obj)
        {
            var separator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            string val = obj != null ? obj.ToString() : string.Empty;
            if (separator == ",")
            {
                val = val.Replace(".", ",");
            }
            return val;
        }
    }
}
