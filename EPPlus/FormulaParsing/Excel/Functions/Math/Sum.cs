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
                    retVal += Calculate(arg, context);                    
                }
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double Calculate(FunctionArgument arg, ParsingContext context)
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
                    retVal += Calculate(item, context);
                }
            }
            else if (arg.Value is ExcelDataProvider.ICellInfo)
            {
                foreach (var c in (ExcelDataProvider.ICellInfo)arg.Value)
                {
                    if (ShouldIgnore(c, context) == false)
                    {
                        retVal += GetNumeric(c.Value);
                    }
                }
            }
            else
            {
                retVal += GetNumeric(arg.Value);
            }
            return retVal;
        }
    }
}
