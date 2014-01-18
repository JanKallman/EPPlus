/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionGraphBuilder :IExpressionGraphBuilder
    {
        private readonly ExpressionGraph _graph = new ExpressionGraph();
        private readonly IExpressionFactory _expressionFactory;
        private readonly ParsingContext _parsingContext;
        private int _tokenIndex = 0;
        private bool _negateNextExpression;

        public ExpressionGraphBuilder(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
            : this(new ExpressionFactory(excelDataProvider, parsingContext), parsingContext)
        {

        }

        public ExpressionGraphBuilder(IExpressionFactory expressionFactory, ParsingContext parsingContext)
        {
            _expressionFactory = expressionFactory;
            _parsingContext = parsingContext;
        }

        public ExpressionGraph Build(IEnumerable<Token> tokens)
        {
            _tokenIndex = 0;
            _graph.Reset();
            BuildUp(tokens, null);
            return _graph;
        }

        private void BuildUp(IEnumerable<Token> tokens, Expression parent)
        {
            while (_tokenIndex < tokens.Count())
            {
                var token = tokens.ElementAt(_tokenIndex);
                IOperator op = null;
                if (token.TokenType == TokenType.Operator && OperatorsDict.Instance.TryGetValue(token.Value, out op))
                {
                    SetOperatorOnExpression(parent, op);
                }
                else if (token.TokenType == TokenType.Function)
                {
                    BuildFunctionExpression(tokens, parent, token.Value);
                }
                else if (token.TokenType == TokenType.OpeningEnumerable)
                {
                    _tokenIndex++;
                    BuildEnumerableExpression(tokens, parent);
                }
                else if (token.TokenType == TokenType.OpeningParenthesis)
                {
                    _tokenIndex++;
                    BuildGroupExpression(tokens, parent);
                    if (parent is FunctionExpression)
                    {
                        return;
                    }
                }
                else if (token.TokenType == TokenType.ClosingParenthesis || token.TokenType == TokenType.ClosingEnumerable)
                {
                    break;
                }
                else if (token.TokenType == TokenType.Negator)
                {
                    _negateNextExpression = true;
                }
                else if(token.TokenType == TokenType.Percent)
                {
                    SetOperatorOnExpression(parent, Operator.Percent);
                    if (parent == null)
                    {
                        _graph.Add(ConstantExpressions.Percent);
                    }
                    else
                    {
                        parent.AddChild(ConstantExpressions.Percent);
                    }
                }
                else
                {
                    CreateAndAppendExpression(ref parent, token);
                }
                _tokenIndex++;
            }
        }

        private void BuildEnumerableExpression(IEnumerable<Token> tokens, Expression parent)
        {
            if (parent == null)
            {
                _graph.Add(new EnumerableExpression());
                BuildUp(tokens, _graph.Current);
            }
            else
            {
                var enumerableExpression = new EnumerableExpression();
                parent.AddChild(enumerableExpression);
                BuildUp(tokens, enumerableExpression);
            }
        }

        private void CreateAndAppendExpression(ref Expression parent, Token token)
        {
            if (IsWaste(token)) return;
            if (parent != null && 
                (token.TokenType == TokenType.Comma || token.TokenType == TokenType.SemiColon))
            {
                parent = parent.PrepareForNextChild();
                return;
            }
            if (_negateNextExpression)
            {
                token.Negate();
                _negateNextExpression = false;
            }
            var expression = _expressionFactory.Create(token);
            if (parent == null)
            {
                _graph.Add(expression);
            }
            else
            {
                parent.AddChild(expression);
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

        private void BuildFunctionExpression(IEnumerable<Token> tokens, Expression parent, string funcName)
        {
            if (parent == null)
            {
                _graph.Add(new FunctionExpression(funcName, _parsingContext));
                HandleFunctionArguments(tokens, _graph.Current);
            }
            else
            {
                var func = new FunctionExpression(funcName, _parsingContext);
                parent.AddChild(func);
                HandleFunctionArguments(tokens, func);
            }
        }

        private void HandleFunctionArguments(IEnumerable<Token> tokens, Expression function)
        {
            _tokenIndex++;
            var token = tokens.ElementAt(_tokenIndex);
            if (token.TokenType != TokenType.OpeningParenthesis)
            {
                throw new ExcelErrorValueException(eErrorType.Value);
            }
            var argExpression = function.AddChild(new FunctionArgumentExpression(function));
            _tokenIndex++;
            BuildUp(tokens, argExpression);
        }

        private void BuildGroupExpression(IEnumerable<Token> tokens, Expression parent)
        {
            if (parent == null)
            {
                _graph.Add(new GroupExpression());
                BuildUp(tokens, _graph.Current);
            }
            else
            {
                if (parent.IsGroupedExpression)
                {
                    var newGroupExpression = new GroupExpression();
                    parent.AddChild(newGroupExpression);
                    BuildUp(tokens, newGroupExpression);
                }
                 BuildUp(tokens, parent);
            }
        }

        private void SetOperatorOnExpression(Expression parent, IOperator op)
        {
            if (parent == null)
            {
                _graph.Current.Operator = op;
            }
            else
            {
                Expression candidate;
                if (parent is FunctionArgumentExpression)
                {
                    candidate = parent.Children.Last();
                }
                else
                {
                    candidate = parent.Children.Last();
                    if (candidate is FunctionArgumentExpression)
                    {
                        candidate = candidate.Children.Last();
                    }
                }
                candidate.Operator = op;
            }
        }
    }
}
