using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    public class ErrorHandlingFunctionCompiler : FunctionCompiler
    {
        public ErrorHandlingFunctionCompiler(ExcelFunction function)
            : base(function)
        {

        }
        public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
        {
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(context);
            foreach (var child in children)
            {
                CompileResult arg = default(CompileResult);
                try
                {
                    arg = child.Compile();
                    BuildFunctionArguments(arg != null ? arg.Result : null, args);
                }
                catch (ExcelFunctionException efe)
                {
                    return ((ErrorHandlingFunction)Function).HandleError(efe.ErrorCode);
                }
                catch (Exception)
                {
                    return ((ErrorHandlingFunction)Function).HandleError(ExcelErrorCodes.Value.Code);
                }
                
            }
            return Function.Execute(args, context);
        }
    }
}
