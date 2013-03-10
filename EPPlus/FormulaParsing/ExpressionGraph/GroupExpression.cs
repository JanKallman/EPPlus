using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class GroupExpression : Expression
    {
        public GroupExpression()
            : this(new ExpressionCompiler())
        {

        }

        public GroupExpression(IExpressionCompiler expressionCompiler)
        {
            _expressionCompiler = expressionCompiler;
        }

        private readonly IExpressionCompiler _expressionCompiler;

        public override CompileResult Compile()
        {
            return _expressionCompiler.Compile(Children);
        }

        public override bool IsGroupedExpression
        {
            get { return true; }
        }
    }
}
