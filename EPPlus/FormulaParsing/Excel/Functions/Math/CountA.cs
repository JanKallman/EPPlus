using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class CountA : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var nItems = 0d;
            Calculate(arguments, context,  ref nItems);
            return CreateResult(nItems, DataType.Integer);
        }

        private void Calculate(IEnumerable<FunctionArgument> items, ParsingContext context, ref double nItems)
        {
            foreach (var item in items)
            {
                if (item.Value is ExcelDataProvider.ICellInfo)
                {
                    foreach (var c in (ExcelDataProvider.ICellInfo)item.Value)
                    {
                        if (ShouldIgnore(c, context) == false && ShouldCount(c.Value))
                        {
                            nItems++;
                        }
                    }
                }
                else if (item.Value is IEnumerable<FunctionArgument>)
                {
                    Calculate((IEnumerable<FunctionArgument>)item.Value, context, ref nItems);
                }
                else if (ShouldCount(item.Value))
                {
                    nItems++;
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
            return (!string.IsNullOrEmpty(value.ToString()));
        }
    }
}
