using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy
{
    public abstract class CompileStrategy
    {
        protected readonly Expression _expression;

        public CompileStrategy(Expression expression)
        {
            _expression = expression;
        }

        public abstract Expression Compile();
    }
}
