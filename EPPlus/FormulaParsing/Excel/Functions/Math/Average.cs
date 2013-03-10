using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Average : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            double nValues = 0d, result = 0d;
            foreach (var arg in arguments)
            {
                Calculate(arg, ref result, ref nValues);
            }
            return CreateResult(result/nValues, DataType.Decimal);
        }

        private void Calculate(FunctionArgument arg, ref double retVal, ref double nValues)
        {
            if (ShouldIgnore(arg))
            {
                return;
            }
            if (arg.Value is double)
            {
                nValues++;
                retVal += Convert.ToDouble(arg.Value);
            }
            else if (arg.Value is int)
            {
                nValues++;
                retVal += Convert.ToDouble((int)arg.Value);
            }
            else if(arg.Value is bool)
            {
                nValues++;
                retVal += (bool)arg.Value ? 1 : 0;
            }
            else if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    Calculate(item, ref retVal, ref nValues);
                }
            }
        }


    }
}
