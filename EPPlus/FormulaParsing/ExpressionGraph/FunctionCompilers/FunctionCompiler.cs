using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System.Collections;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    public abstract class FunctionCompiler
    {
        protected ExcelFunction Function
        {
            get;
            private set;
        }

        public FunctionCompiler(ExcelFunction function)
        {
            Require.That(function).Named("function").IsNotNull();
            Function = function;
        }

        protected void BuildFunctionArguments(object result, List<FunctionArgument> args)
        {
            if (result is IEnumerable<object>)
            {
                var argList = new List<FunctionArgument>();
                foreach (var arg in ((IEnumerable<object>)result))
                {
                    BuildFunctionArguments(arg, argList);
                }
                args.Add(new FunctionArgument(argList));
            }
            else
            {
                args.Add(new FunctionArgument(result));
            }
        }

        public abstract CompileResult Compile(IEnumerable<Expression> children, ParsingContext context);
    }
}
