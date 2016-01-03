using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    public class IfNaFunctionCompiler : FunctionCompiler
    {
        public IfNaFunctionCompiler(ExcelFunction function)
            :base(function)
        {
            
        }

        public override CompileResult Compile(IEnumerable<Expression> children, ParsingContext context)
        {
            if (children.Count() != 2) throw new ExcelErrorValueException(eErrorType.Value);
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(context);
            var firstChild = children.First();
            var lastChild = children.ElementAt(1);
            try
            {
                var result = firstChild.Compile();
                if (result.DataType == DataType.ExcelError && (Equals(result.Result,
                    ExcelErrorValue.Create(eErrorType.NA))))
                {
                    args.Add(new FunctionArgument(lastChild.Compile().Result));
                }
                else
                {
                    args.Add(new FunctionArgument(result.Result));
                }

            }
            catch (ExcelErrorValueException)
            {
                args.Add(new FunctionArgument(lastChild.Compile().Result));
            }
            return Function.Execute(args, context);
        }
    }
}
