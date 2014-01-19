using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionGraphBuilder2 : IExpressionGraphBuilder
    {
        private readonly ExpressionGraph _graph = new ExpressionGraph();
        private readonly List<Expression> _expressions = new List<Expression>();  
        private readonly IExpressionFactory _expressionFactory;
        private readonly ParsingContext _parsingContext;
        private int _tokenIndex = 0;
        private bool _negateNextExpression;

        public ExpressionGraphBuilder2(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
            : this(new ExpressionFactory(excelDataProvider, parsingContext), parsingContext)
        {

        }

        public ExpressionGraphBuilder2(IExpressionFactory expressionFactory, ParsingContext parsingContext)
        {
            _expressionFactory = expressionFactory;
            _parsingContext = parsingContext;
        }

        public ExpressionGraph Build(IEnumerable<Token> tokens)
        {
            _tokenIndex = -1;
            _graph.Reset();
            BuildUp(tokens.ToArray(), null);
            return _graph;
        }

        private Expression _last;

        private void BuildUp(Token[] tokens, Expression parent)
        {
            while (_tokenIndex < tokens.Length - 1)
            {
                _tokenIndex++;
                var token = tokens.ElementAt(_tokenIndex);
                IOperator op = null;
                if (token.TokenType == TokenType.Operator && OperatorsDict.Instance.TryGetValue(token.Value, out op))
                {
                    _last.Operator = op;
                }
                else if (token.TokenType == TokenType.Function)
                {
                    var expression = new FunctionExpression(token.Value, _parsingContext);
                    _last = expression;
                    if (parent != null)
                    {
                        parent.AddChild(expression);
                    }
                    else
                    {
                        _expressions.Add(expression);
                    }
                    BuildUp(tokens, expression);
                }
                else if (token.TokenType == TokenType.OpeningEnumerable)
                {
                    var expression = new EnumerableExpression();
                    _last = expression;
                    if (parent != null)
                    {
                        parent.AddChild(expression);
                    }
                    else
                    {
                        _expressions.Add(expression);
                    }
                    BuildUp(tokens, expression);
                }
                else if (token.TokenType == TokenType.OpeningParenthesis)
                {
                    if (_last is FunctionExpression)
                    {
                        BuildUp(tokens, _last);
                    }
                    else
                    {
                        var expression = new GroupExpression();
                        _last = expression;
                        if (parent != null)
                        {
                            parent.AddChild(expression);
                        }
                        else
                        {
                            _expressions.Add(expression);
                        }
                         BuildUp(tokens, expression);
                    }
                }
                else if (token.TokenType == TokenType.ClosingParenthesis ||
                         token.TokenType == TokenType.ClosingEnumerable)
                {
                    return;
                }
                else if (token.TokenType == TokenType.Negator)
                {
                    _negateNextExpression = true;
                }
                else if (token.TokenType == TokenType.Percent)
                {
                    // TODO: Add Constant expression (0.01) and an muliplying operator
                    // with lower precedence than multiply
                    _last.Operator = Operator.Percent;
                    if (parent == null)
                    {
                        _expressions.Add(ConstantExpressions.Percent);
                    }
                    else
                    {
                        parent.AddChild(ConstantExpressions.Percent);
                    }
                }
                else
                {
                    CreateAndAppendExpression(parent, token);
                }
            }
        }

        private bool IsWaste(Token token)
        {
            if (token.TokenType == TokenType.String)
            {
                return true;
            }
            return false;
        }

        private void CreateAndAppendExpression(Expression parent, Token token)
        {
            if (IsWaste(token)) return;
            if (parent != null &&
                (token.TokenType == TokenType.Comma || token.TokenType == TokenType.SemiColon))
            {
                parent.PrepareForNextChild();
                return;
            }
            if (_negateNextExpression)
            {
                token.Negate();
                _negateNextExpression = false;
            }
            var expression = _expressionFactory.Create(token);
            _last = expression;
            if (parent == null)
            {
                _expressions.Add(expression);
            }
            else
            {
                parent.AddChild(expression);
            }
        }

        #region Old code
        //private void BuildUp(IEnumerable<Token> tokens, Expression parent)
        //{
        //    while (_tokenIndex < tokens.Count())
        //    {
        //        var token = tokens.ElementAt(_tokenIndex);
        //        IOperator op = null;
        //        if (token.TokenType == TokenType.Operator && OperatorsDict.Instance.TryGetValue(token.Value, out op))
        //        {
        //            SetOperatorOnExpression(parent, op);
        //        }
        //        else if (token.TokenType == TokenType.Function)
        //        {
        //            _tokenIndex++;
        //            BuildFunctionExpression(tokens, parent, token.Value);
        //        }
        //        else if (token.TokenType == TokenType.OpeningEnumerable)
        //        {
        //            _tokenIndex++;
        //            BuildEnumerableExpression(tokens, parent);
        //        }
        //        else if (token.TokenType == TokenType.OpeningParenthesis)
        //        {
        //            _tokenIndex++;
        //            BuildGroupExpression(tokens, parent);
        //            if (parent is FunctionExpression)
        //            {
        //                return;
        //            }
        //        }
        //        else if (token.TokenType == TokenType.ClosingParenthesis || token.TokenType == TokenType.ClosingEnumerable)
        //        {
        //            break;
        //        }
        //        else if (token.TokenType == TokenType.Negator)
        //        {
        //            _negateNextExpression = true;
        //        }
        //        else if(token.TokenType == TokenType.Percent)
        //        {
        //            // TODO: Add Constant expression (0.01) and an muliplying operator
        //            // with lower precedence than multiply
        //            SetOperatorOnExpression(parent, Operator.Percent);
        //            if (parent == null)
        //            {
        //                _graph.Add(ConstantExpressions.Percent);
        //            }
        //            else
        //            {
        //                parent.AddChild(ConstantExpressions.Percent);
        //            }
        //        }
        //        else
        //        {
        //            CreateAndAppendExpression(parent, token);
        //        }
        //        _tokenIndex++;
        //    }
        //}

        //private void BuildEnumerableExpression(IEnumerable<Token> tokens, Expression parent)
        //{
        //    if (parent == null)
        //    {
        //        _graph.Add(new EnumerableExpression());
        //        BuildUp(tokens, _graph.Current);
        //    }
        //    else
        //    {
        //        var enumerableExpression = new EnumerableExpression();
        //        parent.AddChild(enumerableExpression);
        //        BuildUp(tokens, enumerableExpression);
        //    }
        //}

        //private void CreateAndAppendExpression(Expression parent, Token token)
        //{
        //    if (IsWaste(token)) return;
        //    if (parent != null && 
        //        (token.TokenType == TokenType.Comma || token.TokenType == TokenType.SemiColon))
        //    {
        //        parent.PrepareForNextChild();
        //        return;
        //    }
        //    if (_negateNextExpression)
        //    {
        //        token.Negate();
        //        _negateNextExpression = false;
        //    }
        //    var expression = _expressionFactory.Create(token);
        //    if (parent == null)
        //    {
        //        _graph.Add(expression);
        //    }
        //    else
        //    {
        //        parent.AddChild(expression);
        //    }
        //}

        //private bool IsWaste(Token token)
        //{
        //    if (token.TokenType == TokenType.String)
        //    {
        //        return true;
        //    }
        //    return false;
        //}

        //private void BuildFunctionExpression(IEnumerable<Token> tokens, Expression parent, string funcName)
        //{
        //    if (parent == null)
        //    {
        //        _graph.Add(new FunctionExpression(funcName, _parsingContext));
        //        BuildUp(tokens, _graph.Current);
        //    }
        //    else
        //    {
        //        var func = new FunctionExpression(funcName, _parsingContext);
        //        parent.AddChild(func);
        //        BuildUp(tokens, func);
        //    }
        //}

        //private void BuildGroupExpression(IEnumerable<Token> tokens, Expression parent)
        //{
        //    if (parent == null)
        //    {
        //        _graph.Add(new GroupExpression());
        //        BuildUp(tokens, _graph.Current);
        //    }
        //    else
        //    {
        //        if (parent.IsGroupedExpression)
        //        {
        //            var newGroupExpression = new GroupExpression();
        //            parent.AddChild(newGroupExpression);
        //            BuildUp(tokens, newGroupExpression);
        //        }
        //        BuildUp(tokens, parent);
        //    }
        //}

        //private void SetOperatorOnExpression(Expression parent, IOperator op)
        //{
        //    if (parent == null)
        //    {
        //        _graph.Current.Operator = op;
        //    }
        //    else
        //    {
        //        var candidate = parent.Children.Last();
        //        if (candidate is FunctionArgumentExpression)
        //        {
        //            candidate = candidate.Children.Last();
        //        }
        //        candidate.Operator = op;
        //    }
        //}
        #endregion
    }
}
