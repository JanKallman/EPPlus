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
using System.Threading;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class TokenSeparatorProvider : ITokenSeparatorProvider
    {
       private static readonly Dictionary<string, Token> _tokens;

        static TokenSeparatorProvider()
        {
            _tokens = new Dictionary<string, Token>();
            _tokens.Add("+", new Token("+", TokenType.Operator));
            _tokens.Add("-", new Token("-", TokenType.Operator));
            _tokens.Add("*", new Token("*", TokenType.Operator));
            _tokens.Add("/", new Token("/", TokenType.Operator));
            _tokens.Add("^", new Token("^", TokenType.Operator));
            _tokens.Add("&", new Token("&", TokenType.Operator));
            _tokens.Add(">", new Token(">", TokenType.Operator));
            _tokens.Add("<", new Token("<", TokenType.Operator));
            _tokens.Add("=", new Token("=", TokenType.Operator));
            _tokens.Add("<=", new Token("<=", TokenType.Operator));
            _tokens.Add(">=", new Token(">=", TokenType.Operator));
            _tokens.Add("<>", new Token("<>", TokenType.Operator));
            _tokens.Add("(", new Token("(", TokenType.OpeningParenthesis));
            _tokens.Add(")", new Token(")", TokenType.ClosingParenthesis));
            _tokens.Add("{", new Token("{", TokenType.OpeningEnumerable));
            _tokens.Add("}", new Token("}", TokenType.ClosingEnumerable));
            _tokens.Add("'", new Token("'", TokenType.String));
            _tokens.Add("\"", new Token("\"", TokenType.String));
            _tokens.Add(",", new Token(",", TokenType.Comma));
            _tokens.Add(";", new Token(";", TokenType.SemiColon));
            _tokens.Add("[", new Token("[", TokenType.OpeningBracket));
            _tokens.Add("]", new Token("]", TokenType.ClosingBracket));
            _tokens.Add("%", new Token("%", TokenType.Percent));
        }

        IDictionary<string, Token> ITokenSeparatorProvider.Tokens
        {
            get { return _tokens; }
        }

        public bool IsOperator(string item)
        {
            Token token;
            if (_tokens.TryGetValue(item, out token))
            {
                if (token.TokenType == TokenType.Operator)
                {
                    return true;
                }
            }
            return false;
        }

        public bool IsPossibleLastPartOfMultipleCharOperator(string part)
        {
            return part == "=" || part == ">";
        }
    }
}
