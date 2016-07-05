using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Rank : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var number = ArgToDecimal(arguments, 0);
            var refer = arguments.ElementAt(1);
            bool asc = false;
            if (arguments.Count() > 2)
            {
                asc = base.ArgToBool(arguments, 2);
            }
            var l = new List<double>();

            foreach (var c in refer.ValueAsRangeInfo)
            {
                var v = Utils.ConvertUtil.GetValueDouble(c.Value, false, true);
                if (!double.IsNaN(v))
                {
                    l.Add(v);
                }
            }
            l.Sort();
            int ix;
            if (asc)
            {
                ix = l.IndexOf(number)+1;
            }
            else
            {
                ix = l.Count-l.LastIndexOf(number);
            }
            if (ix <= 0 || ix>l.Count)
            {
                return new CompileResult(ExcelErrorValue.Create(eErrorType.NA), DataType.ExcelError);
            }
            else
            {
                return CreateResult(ix, DataType.Integer);
            }
        }
    }
}
