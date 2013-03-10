using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class BooleanExpression : AtomicExpression
    {
        public BooleanExpression(string expression)
            : base(expression)
        {

        }
        public override CompileResult Compile()
        {
            return new CompileResult(bool.Parse(ExpressionString), DataType.Boolean);
        }
    }
}
