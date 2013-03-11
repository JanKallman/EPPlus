using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class DateExpression : AtomicExpression
    {
        public DateExpression(string expression)
            : base(expression)
        {

        }

        public override CompileResult Compile()
        {
            var date = double.Parse(ExpressionString,CultureInfo.InvariantCulture);
            return new CompileResult(DateTime.FromOADate(date), DataType.Date);
        }
    }
}
