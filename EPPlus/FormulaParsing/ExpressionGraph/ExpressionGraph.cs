using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionGraph
    {
        private List<Expression> _expressions = new List<Expression>();
        public IEnumerable<Expression> Expressions { get { return _expressions; } }
        public Expression Current { get; private set; }

        public void Add(Expression expression)
        {
            _expressions.Add(expression);
            if (Current != null)
            {
                Current.Next = expression;
                expression.Prev = Current;
            }
            Current = expression;
        }

        public void Reset()
        {
            _expressions.Clear();
            Current = null;
        }

        public void Remove(Expression item)
        {
            if (item == Current)
            {
                Current = item.Prev ?? item.Next;
            }
            _expressions.Remove(item);
        }
    }
}
