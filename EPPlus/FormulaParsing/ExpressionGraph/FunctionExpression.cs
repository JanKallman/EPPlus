using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class FunctionExpression : AtomicExpression
    {
        public FunctionExpression(string expression, ParsingContext parsingContext)
            : base(expression)
        {
            _parsingContext = parsingContext;
        }

        private readonly ParsingContext _parsingContext;
        private readonly FunctionCompilerFactory _functionCompilerFactory = new FunctionCompilerFactory();

        public override CompileResult Compile()
        {
            var function = _parsingContext.Configuration.FunctionRepository.GetFunction(ExpressionString);
            var compiler = _functionCompilerFactory.Create(function);
            return compiler.Compile(Children, _parsingContext);
        }

        public override void PrepareForNextChild()
        {
            base.AddChild(new FunctionArgumentExpression());
        }

        public override Expression AddChild(Expression child)
        {
            if (Children.Count() == 0)
            {
                var group = base.AddChild(new FunctionArgumentExpression());
                group.AddChild(child);
            }
            else
            {
                Children.Last().AddChild(child);
            }
            return child;
        }

        public override Expression MergeWithNext()
        {
            Expression returnValue = null;
            if (Next != null && Operator != null)
            {
                var result = Operator.Apply(Compile(), Next.Compile());
                var expressionString = result.Result.ToString();
                var converter = new ExpressionConverter();
                returnValue = converter.FromCompileResult(result);
                if (Next != null)
                {
                    Operator = Next.Operator;
                }
                else
                {
                    Operator = null;
                }
                Next = Next.Next;
            }
            return returnValue;
        }
    }
}
