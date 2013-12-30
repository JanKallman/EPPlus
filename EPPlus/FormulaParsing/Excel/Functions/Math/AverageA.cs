using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class AverageA : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1, eErrorType.Div0);
            double nValues = 0d, result = 0d;
            foreach (var arg in arguments)
            {
                Calculate(arg, context, ref result, ref nValues);
            }
            return CreateResult(result / nValues, DataType.Decimal);
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
            else if (arg.IsExcelRange)
            {
                foreach (var c in arg.ValueAsRangeInfo)
                {
                    if (ShouldIgnore(c, context)) continue;
                    CheckForAndHandleExcelError(c);
                    if (IsNumber(c.Value))
                    {
                        nValues++;
                        retVal += c.ValueDouble;
                    }
                    else if (c.Value is bool)
                    {
                        nValues++;
                        retVal += (bool) c.Value ? 1 : 0;
                    }
                    else if (c.Value is string)
                    {
                        nValues++;
                    }
                }
            }
            else
            {
                var numericValue = GetNumericValue(arg.Value);
                if (numericValue.HasValue)
                {
                    nValues++;
                    retVal += numericValue.Value;
                }
                else if((arg.Value is string) && !ConvertUtil.IsNumericString(arg.Value))
                {
                    ThrowExcelErrorValueException(eErrorType.Value);
                }
            }
            CheckForAndHandleExcelError(arg);
        }

        private double? GetNumericValue(object obj)
        {
            if (IsNumber(obj) || (obj is bool))
            {
                return ConvertUtil.GetValueDouble(obj);
            }
            else if (ConvertUtil.IsNumericString(obj))
            {
                return double.Parse(obj.ToString(), CultureInfo.InvariantCulture);
            }
            return default(double?);
        }
    }
}
