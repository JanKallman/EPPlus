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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class SourceCodeTokenizer : ISourceCodeTokenizer
    {
        public static ISourceCodeTokenizer Default
        {
            get { return new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty); }
        }

        public SourceCodeTokenizer(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider)
            : this(new TokenFactory(functionRepository, nameValueProvider), new TokenSeparatorProvider())
        {

        }
        public SourceCodeTokenizer(ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider)
        {
            _tokenFactory = tokenFactory;
            _tokenProvider = tokenProvider;
        }

        private readonly ITokenSeparatorProvider _tokenProvider;
        private readonly ITokenFactory _tokenFactory;

        public IEnumerable<Token> Tokenize(string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return Enumerable.Empty<Token>();
            }
            var context = new TokenizerContext(input);
            foreach (var c in context.FormulaChars)
            {
                Token tokenSeparator;
                if(CharIsTokenSeparator(c, out tokenSeparator))
                {
                    if (context.IsInString && tokenSeparator.TokenType != TokenType.String)
                    {
                        context.AppendToCurrentToken(c);
                        continue;
                    }
                    if (tokenSeparator.TokenType == TokenType.OpeningBracket)
                    {
                        context.AppendToCurrentToken(c);
                        context.BracketCount++;
                        continue;
                    }
                    if (tokenSeparator.TokenType == TokenType.ClosingBracket)
                    {
                        context.AppendToCurrentToken(c);
                        context.BracketCount--;
                        continue;
                    }
                    if (context.BracketCount > 0)
                    {
                        context.AppendToCurrentToken(c);
                        continue;
                    }
                    // two operators in sequence could be "<=" or ">="
                    if (IsPartOfMultipleCharSeparator(context, c))
                    {
                        context.AppendToLastToken(c.ToString());
                        context.NewToken();
                        continue;
                    }
                    if (tokenSeparator.TokenType == TokenType.String)
                    {
                        if (context.LastToken != null &&
                            context.LastToken.TokenType == TokenType.String && 
                            !context.CurrentTokenHasValue)
                        {
                            // We are dealing with an empty string ('').
                            context.AddToken(new Token(string.Empty, TokenType.StringContent));
                        }
                        context.ToggleIsInString();
                    }
                    if (context.CurrentTokenHasValue)
                    {
                        context.AddToken(CreateToken(context));
                    }
                    if (tokenSeparator.Value == "-")
                    {
                        if (TokenIsNegator(context))
                        {
                            context.AddToken(new Token("-", TokenType.Negator));
                            continue;
                        }
                    }
                    context.AddToken(tokenSeparator);
                    context.NewToken();
                    continue;
                }
                context.AppendToCurrentToken(c);
            }
            if (context.CurrentTokenHasValue)
            {
                context.AddToken(CreateToken(context));
            }
            return context.Result;
        }

        private static bool TokenIsNegator(TokenizerContext context)
        {
            return context.LastToken == null
                                        ||
                                        context.LastToken.TokenType == TokenType.Operator
                                        ||
                                        context.LastToken.TokenType == TokenType.OpeningParenthesis
                                        ||
                                        context.LastToken.TokenType == TokenType.Comma
                                        ||
                                        context.LastToken.TokenType == TokenType.OpeningEnumerable;
        }

        private bool IsPartOfMultipleCharSeparator(TokenizerContext context, char c)
        {
            var lastToken = context.LastToken != null ? context.LastToken.Value : string.Empty;
            return _tokenProvider.IsOperator(lastToken) 
                && _tokenProvider.IsPossibleLastPartOfMultipleCharOperator(c.ToString()) 
                && !context.CurrentTokenHasValue;
        }

        private Token CreateToken(TokenizerContext context)
        {
            if (context.CurrentToken == "-")
            {
                if (context.LastToken == null && context.LastToken.TokenType == TokenType.Operator)
                {
                    return new Token("-", TokenType.Negator);
                }
            }
            return _tokenFactory.Create(context.Result, context.CurrentToken);
        }

        private bool CharIsTokenSeparator(char c, out Token token)
        {
            var result = _tokenProvider.Tokens.ContainsKey(c.ToString());
            token = result ? token = _tokenProvider.Tokens[c.ToString()] : null;
            return result;
        }
    }
}
