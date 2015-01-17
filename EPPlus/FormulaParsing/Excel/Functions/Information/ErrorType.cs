using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    public class ErrorType : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var error = arguments.ElementAt(0);
            var isErrorFunc = context.Configuration.FunctionRepository.GetFunction("iserror");
            var isErrorResult = isErrorFunc.Execute(arguments, context);
            if (!(bool) isErrorResult.Result)
            {
                return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
            }
            var errorType = error.ValueAsExcelErrorValue;
            int retValue;
            switch (errorType.Type)
            {
                case eErrorType.Null:
                    return CreateResult(1, DataType.Integer);
                case eErrorType.Div0:
                    return CreateResult(2, DataType.Integer);
                case eErrorType.Value:
                    return CreateResult(3, DataType.Integer);
                case eErrorType.Ref:
                    return CreateResult(4, DataType.Integer);
                case eErrorType.Name:
                    return CreateResult(5, DataType.Integer);
                case eErrorType.Num:
                    return CreateResult(6, DataType.Integer);
                case eErrorType.NA:
                    return CreateResult(7, DataType.Integer);
            }
            return CreateResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
        }
    }
}
