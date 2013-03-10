using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy
{
    public class CompileStrategyFactory : ICompileStrategyFactory
    {
        public CompileStrategy Create(Expression expression)
        {
            if (expression.Operator.Operator == Operators.Concat)
            {
                return new StringConcatStrategy(expression);
            }
            else
            {
                return new DefaultCompileStrategy(expression);
            }
        }
    }
}
