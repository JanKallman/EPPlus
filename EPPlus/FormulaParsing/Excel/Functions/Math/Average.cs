using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

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
                Calculate(arg, context, ref result, ref nValues);
            }
            return CreateResult(result/nValues, DataType.Decimal);
        }

        private void Calculate(FunctionArgument arg, ParsingContext context, ref double retVal, ref double nValues)
        {
            if (ShouldIgnore(arg))
            {
                return;
            }
            if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    Calculate(item, context, ref retVal, ref nValues);
                }
            }
            else if (arg.Value is ExcelDataProvider.IRangeInfo)
            {
                foreach (var c in (ExcelDataProvider.IRangeInfo)arg.Value)
                {
                    if (!ShouldIgnore(c, context))
                    {
                        nValues++;
                        retVal += (double)c.ValueDouble;
                    }
                }
            } 
            else if (IsNumeric(arg.Value))
            {
                nValues++;
                retVal += ConvertUtil.GetValueDouble(arg.Value, true);
            }
            //else if (arg.Value is int)
            //{
            //    nValues++;
            //    retVal += Convert.ToDouble((int)arg.Value);
            //}
            //else if(arg.Value is bool)
            //{
            //    nValues++;
            //    retVal += (bool)arg.Value ? 1 : 0;
            //}
            
        }


    }
}
