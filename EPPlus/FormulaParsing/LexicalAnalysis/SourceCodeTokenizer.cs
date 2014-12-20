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
            _tokenProvider = tokenProvider;
        }

        private readonly ITokenSeparatorProvider _tokenProvider;
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
            for (int i = 0; i<context.FormulaChars.Length;i++)
            {
                var c = context.FormulaChars[i];
                Token tokenSeparator;
                if (CharIsTokenSeparator(c, out tokenSeparator))
                {
                    if (context.IsInString)
                    {
                        if (IsDoubleQuote(tokenSeparator, i, context))
                        {
                            i ++;
                            context.AppendToCurrentToken(c);
                            continue;
                        }
                        if(tokenSeparator.TokenType != TokenType.String)
                        {
                            context.AppendToCurrentToken(c);
                            continue;
                        }
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
                        var sOp = context.LastToken.Value + c.ToString(CultureInfo.InvariantCulture);
                        var op = _tokenProvider.Tokens[sOp];
                        context.ReplaceLastToken(op);
                        context.NewToken();
                        continue;
                    }
                    if (tokenSeparator.TokenType == TokenType.String)
                    {
                        if (context.LastToken != null && context.LastToken.TokenType == TokenType.OpeningEnumerable)
                        {
                            context.AppendToCurrentToken(c);
                            context.ToggleIsInString();
                            continue;
                        }
                        else if (context.LastToken != null &&
                            context.LastToken.TokenType == TokenType.String &&
                            !context.CurrentTokenHasValue) //Added check for enumartion 
                        {
                            // We are dealing with an empty string ('').
                            context.AddToken(new Token(string.Empty, TokenType.StringContent));
                        }
                        context.ToggleIsInString();
                    }
                    if (context.CurrentTokenHasValue)
                    {
                        if (Regex.IsMatch(context.CurrentToken, "^\"*$"))
                        {
                            context.AddToken(_tokenFactory.Create(context.CurrentToken, TokenType.StringContent));
                        }
                        else
                        {
                            context.AddToken(CreateToken(context, worksheet));  
                        }
                        
                        
                        //If the a next token is an opening parantheses and the previous token is interpeted as an address or name, then the currenct token is a function
                        if(tokenSeparator.TokenType==TokenType.OpeningParenthesis && (context.LastToken.TokenType==TokenType.ExcelAddress || context.LastToken.TokenType==TokenType.NameValue)) 
                        {
                            context.LastToken.TokenType=TokenType.Function;
                        }
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
                context.AddToken(CreateToken(context, worksheet));
            }

            CleanupTokens(context);

            return context.Result;
        }

        private static bool IsDoubleQuote(Token tokenSeparator, int formulaCharIndex, TokenizerContext context)
        {
            return tokenSeparator.TokenType == TokenType.String && formulaCharIndex + 1 < context.FormulaChars.Length && context.FormulaChars[formulaCharIndex + 1] == '\"';
        }


        private static void CleanupTokens(TokenizerContext context)
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
                else if ((token.TokenType == TokenType.Operator || token.TokenType == TokenType.Negator) && i < context.Result.Count - 1 &&
                         (token.Value=="+" || token.Value=="-"))
                {
                    if (i > 0 && token.Value == "+")    //Remove any + with an opening parenthesis before.
                    {
                        if (context.Result[i - 1].TokenType  == TokenType.OpeningParenthesis)
                        {
                            context.Result.RemoveAt(i);
                            SetNegatorOperator(context, i);
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
                            SetNegatorOperator(context, i);
                            i--;
                        }
                        else if (token.Value == "-" && nextToken.Value == "+")
                        {
                            //Remove second
                            context.Result.RemoveAt(i+1);
                            SetNegatorOperator(context, i);
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
                                context.Result[i].TokenType = TokenType.Operator;
                                context.Result[i].Value = "+";
                                SetNegatorOperator(context, i);
                                i--;
                            }
                        }
                    }
                }
            }
        }

        private static void SetNegatorOperator(TokenizerContext context, int i)
        {
            if (context.Result[i].Value == "-" && i > 0 && (context.Result[i].TokenType == TokenType.Operator || context.Result[i].TokenType == TokenType.Negator))
            {
                if (TokenIsNegator(context.Result[i - 1]))
                {
                    context.Result[i].TokenType = TokenType.Negator;
                }
                else
                {
                    context.Result[i].TokenType = TokenType.Operator;
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

        private bool IsPartOfMultipleCharSeparator(TokenizerContext context, char c)
        {
            var lastToken = context.LastToken != null ? context.LastToken.Value : string.Empty;
            return _tokenProvider.IsOperator(lastToken) 
                && _tokenProvider.IsPossibleLastPartOfMultipleCharOperator(c.ToString(CultureInfo.InvariantCulture)) 
                && !context.CurrentTokenHasValue;
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

        private bool CharIsTokenSeparator(char c, out Token token)
        {
            var result = _tokenProvider.Tokens.ContainsKey(c.ToString());
            token = result ? token = _tokenProvider.Tokens[c.ToString()] : null;
            return result;
        }
    }
}
