using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionCompiler : IExpressionCompiler
    {
        private IEnumerable<Expression> _expressions;
        private IExpressionConverter _expressionConverter;
        private ICompileStrategyFactory _compileStrategyFactory;

        public ExpressionCompiler()
            : this(new ExpressionConverter(), new CompileStrategyFactory())
        {
 
        }

        public ExpressionCompiler(IExpressionConverter expressionConverter, ICompileStrategyFactory compileStrategyFactory)
        {
            _expressionConverter = expressionConverter;
            _compileStrategyFactory = compileStrategyFactory;
        }

        public CompileResult Compile(IEnumerable<Expression> expressions)
        {
            _expressions = expressions;
            return PerformCompilation();
        }

        private CompileResult PerformCompilation()
        {
            var compiledExpressions = HandleGroupedExpressions();
            while(compiledExpressions.Any(x => x.Operator != null))
            {
                var prec = FindLowestPrecedence();
                compiledExpressions = HandlePrecedenceLevel(prec);
            }
            if (_expressions.Any())
            {
                return compiledExpressions.First().Compile();
            }
            return CompileResult.Empty;
        }

        private IEnumerable<Expression> HandleGroupedExpressions()
        {
            if (!_expressions.Any()) return Enumerable.Empty<Expression>();
            var first = _expressions.First();
            var groupedExpressions = _expressions.Where(x => x.IsGroupedExpression);
            foreach(var groupedExpression in groupedExpressions)
            {
                var result = groupedExpression.Compile();
                if (result == CompileResult.Empty) continue;
                var newExp = _expressionConverter.FromCompileResult(result);
                newExp.Operator = groupedExpression.Operator;
                newExp.Prev = groupedExpression.Prev;
                newExp.Next = groupedExpression.Next;
                if (groupedExpression.Prev != null)
                {
                    groupedExpression.Prev.Next = newExp;
                }
                if (groupedExpression == first)
                {
                    first = newExp;
                }
            }
            return RefreshList(first);
        }

        private IEnumerable<Expression> HandlePrecedenceLevel(int precedence)
        {
            var first = _expressions.First();
            var expressionsToHandle = _expressions.Where(x => x.Operator != null && x.Operator.Precedence == precedence);
            foreach (var expression in expressionsToHandle)
            {
                var isFirst = (expression == first);
                var strategy = _compileStrategyFactory.Create(expression);
                var compiledExpression = strategy.Compile();
                if (expression == first)
                {
                    first = compiledExpression;
                }
            }
            return RefreshList(first);
        }

        private int FindLowestPrecedence()
        {
            return _expressions.Where(x => x.Operator != null).Min(x => x.Operator.Precedence);
        }

        private IEnumerable<Expression> RefreshList(Expression first)
        {
            var resultList = new List<Expression>();
            var exp = first;
            resultList.Add(exp);
            while (exp.Next != null)
            {
                resultList.Add(exp.Next);
                exp = exp.Next;
            }
            _expressions = resultList;
            return resultList;
        }
    }
}
