using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
    public class IfError : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var firstArg = arguments.First();
            return GetResultByObject(firstArg.Value);
        }

        private CompileResult GetResultByObject(object result)
        {
            if (IsNumeric(result))
            {
                return CreateResult(result, DataType.Decimal);
            }
            if (result is string)
            {
                return CreateResult(result, DataType.String);
            }
            if (ExcelErrorValue.Values.IsErrorValue(result))
            {
                return CreateResult(result, DataType.ExcelAddress);
            }
            if (result == null)
            {
                return CompileResult.Empty;
            }
            return CreateResult(result, DataType.Enumerable);
        }
    }
}
