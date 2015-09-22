using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public static class ConstantExpressions
    {
        public static Expression Percent
        {
            get { return new ConstantExpression("Percent", () => new CompileResult(0.01, DataType.Decimal)); }
        }
    }

    public class ConstantExpression : AtomicExpression
    {
        private readonly Func<CompileResult> _factoryMethod;

        public ConstantExpression(string title, Func<CompileResult> factoryMethod)
            : base(title)
        {
            _factoryMethod = factoryMethod;
        }

        public override CompileResult Compile()
        {
            return _factoryMethod();
        }
    }
}
