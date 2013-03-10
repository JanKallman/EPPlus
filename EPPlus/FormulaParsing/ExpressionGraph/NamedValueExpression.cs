using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class NamedValueExpression : AtomicExpression
    {
        public NamedValueExpression(string expression, ParsingContext parsingContext)
            : base(expression)
        {
            _parsingContext = parsingContext;
        }

        private readonly ParsingContext _parsingContext;

        public override CompileResult Compile()
        {
            var value = _parsingContext.NameValueProvider.GetNamedValue(ExpressionString);
            var result = _parsingContext.Parser.Parse(value.ToString());
            return new CompileResultFactory().Create(result);
        }
    }
}
