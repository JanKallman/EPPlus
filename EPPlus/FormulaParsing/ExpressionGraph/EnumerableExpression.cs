using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class EnumerableExpression : Expression
    {
        public EnumerableExpression()
            : this(new ExpressionCompiler())
        {

        }

        public EnumerableExpression(IExpressionCompiler expressionCompiler)
        {
            _expressionCompiler = expressionCompiler;
        }

        private readonly IExpressionCompiler _expressionCompiler;

        public override bool IsGroupedExpression
        {
            get { return false; }
        }

        public override void PrepareForNextChild()
        {

        }

        public override CompileResult Compile()
        {
            var result = new List<object>();
            foreach (var childExpression in Children)
            {
                result.Add(_expressionCompiler.Compile(new List<Expression>{ childExpression }).Result);
            }
            return new CompileResult(result, DataType.Enumerable);
        }
    }
}
