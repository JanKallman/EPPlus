using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Count : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var nItems = 0d;
            Calculate(arguments, ref nItems, context);
            return CreateResult(nItems, DataType.Integer);
        }

        private void Calculate(IEnumerable<FunctionArgument> items, ref double nItems, ParsingContext context)
        {
            foreach (var item in items)
            {
                var cs = item.Value as ExcelDataProvider.ICellInfo;
                if (cs != null)
                {
                    foreach (var c in cs)
                    {
                        if (ShouldIgnore(c, context) == false && ShouldCount(c.Value))
                        {
                            nItems++;
                        }
                    }
                }
                else
                {
                    var value = item.Value as IEnumerable<FunctionArgument>;
                    if (value != null)
                    {
                        Calculate(value, ref nItems, context);
                    }
                    else if (ShouldIgnore(item) == false && ShouldCount(item.Value))
                    {
                        nItems++;
                    }
                }
            }
        }

        private bool ShouldCount(object value)
        {
            //if (ShouldIgnore(item))
            //{
            //    return false;
            //}
            if (value == null) return false;
            if (value is int
                ||
                value is double
                ||
                value is decimal
                ||
                value is System.DateTime)
            {
                return true;
            }
            return false;
        }
    }
}
