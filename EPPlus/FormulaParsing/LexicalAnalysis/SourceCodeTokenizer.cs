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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
            _separatorProvider = tokenProvider;
        }

        private readonly ITokenSeparatorProvider _separatorProvider;
        private readonly ITokenFactory _tokenFactory;

        public IEnumerable<Token> Tokenize(string input)
        {
            return Tokenize(input, null);
        }
        public IEnumerable<Token> Tokenize(string input, string worksheet)
        {
            if (string.IsNullOrEmpty(input))
            {
                return Enumerable.Empty<Token>();
            }
            // MA 1401: Ignore leading plus in formula.
            input = input.TrimStart('+');
            var context = new TokenizerContext(input);
            var handler = new TokenHandler(context, _tokenFactory, _separatorProvider);
            handler.Worksheet = worksheet;
            while(handler.HasMore())
            {
                handler.Next();
            }
            if (context.CurrentTokenHasValue)
            {
                context.AddToken(CreateToken(context, worksheet));
            }

            CleanupTokens(context, _separatorProvider.Tokens);

            return context.Result;
        }

        


        private static void CleanupTokens(TokenizerContext context, IDictionary<string, Token>  tokens)
        {
            for (int i = 0; i < context.Result.Count; i++)
            {
                var token=context.Result[i];
                if (token.TokenType == TokenType.Unrecognized)
                {
                    if (i < context.Result.Count - 1)
                    {
                        if (context.Result[i+1].TokenType == TokenType.OpeningParenthesis)
                        {
                            token.TokenType = TokenType.Function;
                        }
                        else
                        {
                            token.TokenType = TokenType.NameValue;
                        }
                    }
                    else
                    {
                        token.TokenType = TokenType.NameValue;
                    }
                }
                else if(token.TokenType == TokenType.WorksheetName){
                    // use this and the following three tokens
                    token.TokenType = TokenType.ExcelAddress;
                    var sb = new StringBuilder();
                    var nToRemove = 3;
                    if (context.Result.Count < i + nToRemove)
                    {
                        token.TokenType = TokenType.InvalidReference;
                        nToRemove = context.Result.Count - i - 1;
                    }
                    else if(context.Result[i + 3].TokenType != TokenType.ExcelAddress)
                    {
                        token.TokenType = TokenType.InvalidReference;
                        nToRemove--;
                    }
                    else
                    {
                        for (var ix = 0; ix < 4; ix++)
                        {
                            sb.Append(context.Result[i + ix].Value);
                        }
                    }
                    token.Value = sb.ToString();
                    for(var ix = 0; ix < nToRemove; ix++)
                    {
                        context.Result.RemoveAt(i + 1);
                    }
                }
                else if ((token.TokenType == TokenType.Operator || token.TokenType == TokenType.Negator) && i < context.Result.Count - 1 &&
                         (token.Value=="+" || token.Value=="-"))
                {
                    if (i > 0 && token.Value == "+")    //Remove any + with an opening parenthesis before.
                    {
                        if (context.Result[i - 1].TokenType  == TokenType.OpeningParenthesis)
                        {
                            context.Result.RemoveAt(i);
                            SetNegatorOperator(context, i, tokens);
                            i--;
                            continue;
                        }
                    }

                    var nextToken = context.Result[i + 1];
                    if (nextToken.TokenType == TokenType.Operator || nextToken.TokenType == TokenType.Negator)
                    {
                        if (token.Value == "+" && (nextToken.Value=="+" || nextToken.Value == "-"))
                        {
                            //Remove first
                            context.Result.RemoveAt(i);
                            SetNegatorOperator(context, i, tokens);
                            i--;
                        }
                        else if (token.Value == "-" && nextToken.Value == "+")
                        {
                            //Remove second
                            context.Result.RemoveAt(i+1);
                            SetNegatorOperator(context, i, tokens);
                            i--;
                        }
                        else if (token.Value == "-" && nextToken.Value == "-")
                        {
                            //Remove first and set operator to +
                            context.Result.RemoveAt(i);
                            if (i == 0)
                            {
                                context.Result.RemoveAt(i+1);
                                i += 2;
                            }
                            else
                            {
                                //context.Result[i].TokenType = TokenType.Operator;
                                //context.Result[i].Value = "+";
                                context.Result[i] = tokens["+"];
                                SetNegatorOperator(context, i, tokens);
                                i--;
                            }
                        }
                    }
                }
            }
        }

        private static void SetNegatorOperator(TokenizerContext context, int i, IDictionary<string, Token>  tokens)
        {
            if (context.Result[i].Value == "-" && i > 0 && (context.Result[i].TokenType == TokenType.Operator || context.Result[i].TokenType == TokenType.Negator))
            {
                if (TokenIsNegator(context.Result[i - 1]))
                {
                    context.Result[i] = new Token("-", TokenType.Negator);
                }
                else
                {
                    context.Result[i] = tokens["-"];
                }
            }
        }

        private static bool TokenIsNegator(TokenizerContext context)
        {
            return TokenIsNegator(context.LastToken);
        }
        private static bool TokenIsNegator(Token t)
        {
            return t == null
                        ||
                        t.TokenType == TokenType.Operator
                        ||
                        t.TokenType == TokenType.OpeningParenthesis
                        ||
                        t.TokenType == TokenType.Comma
                        ||
                        t.TokenType == TokenType.SemiColon
                        ||
                        t.TokenType == TokenType.OpeningEnumerable;
        }

        private Token CreateToken(TokenizerContext context, string worksheet)
        {
            if (context.CurrentToken == "-")
            {
                if (context.LastToken == null && context.LastToken.TokenType == TokenType.Operator)
                {
                    return new Token("-", TokenType.Negator);
                }
            }
            return _tokenFactory.Create(context.Result, context.CurrentToken, worksheet);
        }
    }
}
