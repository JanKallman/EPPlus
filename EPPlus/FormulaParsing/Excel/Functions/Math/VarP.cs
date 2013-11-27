using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class VarP : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var args = ArgsToDoubleEnumerable(arguments, context);
            double avg = args.Average(); 
            double d = args.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2)); 
            var result = d / args.Count(); 
            return new CompileResult(result, DataType.Decimal);
        }
    }
}
