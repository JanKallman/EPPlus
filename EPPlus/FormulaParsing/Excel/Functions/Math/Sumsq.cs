using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Sumsq : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var retVal = 0d;
            if (arguments != null)
            {
                foreach (var arg in arguments)
                {
                    retVal += Calculate(arg, context);
                }
            }
            return CreateResult(retVal, DataType.Decimal);
        }


        private double Calculate(FunctionArgument arg, ParsingContext context, bool isInArray = false)
        {
            var retVal = 0d;
            if (ShouldIgnore(arg))
            {
                return retVal;
            }
            if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    retVal += Calculate(item, context, true);
                }
            }
            else
            {
                var cs = arg.Value as ExcelDataProvider.IRangeInfo;
                if (cs != null)
                {
                    foreach (var c in cs)
                    {
                        if (ShouldIgnore(c, context) == false)
                        {
                            CheckForAndHandleExcelError(c);
                            retVal += System.Math.Pow(c.ValueDouble, 2);
                        }
                    }
                }
                else
                {
                    CheckForAndHandleExcelError(arg);
                    if (IsNumericString(arg.Value) && !isInArray)
                    {
                        var value = ConvertUtil.GetValueDouble(arg.Value);
                        return System.Math.Pow(value, 2);
                    }
                    var ignoreBool = isInArray;
                    retVal += System.Math.Pow(ConvertUtil.GetValueDouble(arg.Value, ignoreBool), 2);
                }
            }
            return retVal;
        }
    }
}
