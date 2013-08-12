using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Sum : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var retVal = 0d;
            if (arguments != null)
            {
                foreach (var arg in arguments)
                {
                    retVal += Calculate(arg);
                }
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double Calculate(FunctionArgument arg)
        {
            var retVal = 0d;
            if (ShouldIgnore(arg))
            {
                return retVal;
            }
            if (arg.Value is double || arg.Value is int)
            {
                retVal += Convert.ToDouble(arg.Value);
            }
            else if (arg.Value is System.DateTime)
            {
                retVal += Convert.ToDateTime(arg.Value).ToOADate();
            }
            else if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    retVal += Calculate(item);
                }
            }
            return retVal;
        }
    }
}
