using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MathObj = System.Math;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class StdevP : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var args = ArgsToDoubleEnumerable(arguments);
            return CreateResult(StandardDeviation(args), DataType.Decimal);
        }

        private static double StandardDeviation(IEnumerable<double> values)
        {
            double avg = values.Average();
            return MathObj.Sqrt(values.Average(v => MathObj.Pow(v - avg, 2)));
        }
    }
}
