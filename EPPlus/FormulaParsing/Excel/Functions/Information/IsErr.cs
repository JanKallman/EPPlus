using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    public class IsErr : ErrorHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var isError = new IsError();
            var result = isError.Execute(arguments, context);
            if ((bool) result.Result)
            {
                var arg = GetFirstValue(arguments);
                if (arg is ExcelDataProvider.IRangeInfo)
                {
                    var r = (ExcelDataProvider.IRangeInfo)arg;
                    var e=r.GetValue(r.Address._fromRow, r.Address._fromCol) as ExcelErrorValue;
                    if (e !=null && e.Type==eErrorType.NA)
                    {
                        return CreateResult(false, DataType.Boolean);
                    }
                }
                else
                {
                    if (arg is ExcelErrorValue && ((ExcelErrorValue)arg).Type==eErrorType.NA)
                    {
                        return CreateResult(false, DataType.Boolean);
                    }
                }
            }
            return result;
        }

        public override CompileResult HandleError(string errorCode)
        {
            return CreateResult(true, DataType.Boolean);
        }
    }
}
