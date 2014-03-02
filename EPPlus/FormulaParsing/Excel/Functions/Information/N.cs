using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    public class N : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var arg = arguments.ElementAt(0);
            
            if (arg.Value is bool)
            {
                var val = (bool) arg.Value ? 1d : 0d;
                return CreateResult(val, DataType.Decimal);
            }
            else if (IsNumeric(arg.Value))
            {
                var val = ConvertUtil.GetValueDouble(arg.Value);
                return CreateResult(val, DataType.Decimal);
            }
            else if (arg.Value is string)
            {
                return CreateResult(0d, DataType.Decimal);
            }
            else if (arg.Value is ExcelErrorValue)
            {
                return CreateResult(arg.Value, DataType.ExcelError);
            }
            throw new ExcelErrorValueException(eErrorType.Value);
        }
    }
}
