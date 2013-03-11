using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class DecimalExpression : AtomicExpression
    {
        public DecimalExpression(string expression)
            : base(expression)
        {
            
        }

        public override CompileResult Compile()
        {
            //Remove JK 2013-03-11. Used CultureInfo.InvariantCulture as an alternative
            //string exp = string.Empty;
            //var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;
            //if (decimalSeparator == ",")
            //{
                
            //    exp = ExpressionString.Replace('.', ',');
            //}
            return new CompileResult(double.Parse(ExpressionString, CultureInfo.InvariantCulture), DataType.Decimal);
        }
    }
}
