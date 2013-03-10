using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy
{
    public class StringConcatStrategy : CompileStrategy
    {
        public StringConcatStrategy(Expression expression)
            : base(expression)
        {
           
        }

        public override Expression Compile()
        {
            var newExp = ExpressionConverter.Instance.ToStringExpression(_expression);
            newExp.Prev = _expression.Prev;
            newExp.Next = _expression.Next;
            if (_expression.Prev != null)
            {
                _expression.Prev.Next = newExp;
            }
            if (_expression.Next != null)
            {
                _expression.Next.Prev = newExp;
            }
            return newExp.MergeWithNext();
        }
    }
}
