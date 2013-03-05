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
 * Jan Källman                      Replaced Adress validate    2013-03-01
 * *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
//using ExcelFormulaParser.Engine.Excel.Functions;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class TokenFactory : ITokenFactory
    {
        public TokenFactory(Dictionary<string, object> functionRepository, Dictionary<string, object> nameValueProvider)
            : this(new TokenSeparatorProvider(), nameValueProvider, functionRepository)
        {

        }

        public TokenFactory(ITokenSeparatorProvider tokenSeparatorProvider, Dictionary<string, object> nameValueProvider, Dictionary<string, object> functionRepository)
        {
            _tokenSeparatorProvider = tokenSeparatorProvider;
            _functionRepository = functionRepository;
            _nameValueProvider = nameValueProvider;
        }

        private readonly ITokenSeparatorProvider _tokenSeparatorProvider;
        private readonly Dictionary<string, object> _functionRepository;
        private readonly Dictionary<string, object> _nameValueProvider;

        public Token Create(IEnumerable<Token> tokens, string token)
        {
            Token tokenSeparator = null;
            if (_tokenSeparatorProvider.Tokens.TryGetValue(token, out tokenSeparator))
            {
                return tokenSeparator;
            }
            if (tokens.Any() && tokens.Last().TokenType == TokenType.String)
            {
                return new Token(token, TokenType.StringContent);
            }
            if (!string.IsNullOrEmpty(token))
            {
                token = token.Trim();
            }
            if (Regex.IsMatch(token, RegexConstants.Decimal))
            {
                return new Token(token, TokenType.Decimal);
            }
            if(Regex.IsMatch(token, RegexConstants.Integer))
            {
                return new Token(token, TokenType.Integer);
            }
            if (Regex.IsMatch(token, RegexConstants.Boolean, RegexOptions.IgnoreCase))
            {
                return new Token(token, TokenType.Boolean);
            }
            if (IsValidAddress(token))
            {
                return new Token(token, TokenType.ExcelAddress);
            }
            if (_functionRepository.ContainsKey(token))
            {
                return new Token(token, TokenType.Function);
            }
            if (_nameValueProvider.ContainsKey(token))
            {
                return new Token(token, TokenType.NameValue);
            }
            return new Token(token, TokenType.Unrecognized);

        }
        readonly char[] SheetNameInvalidChars = new char[] {'?',':','*','/','\\'};
        private bool IsValidAddress(string token)
        {
            int ix;
            if (token[0] == '\'')
            {
                ix = token.LastIndexOf('\'');
                if (ix > 0 && ix < token.Length-1 && token[ix + 1] == '!')
                {
                    if (token.IndexOfAny(SheetNameInvalidChars, 1, ix - 1) > 0)
                    {
                        return false;
                    }
                    token = token.Substring(ix + 2);
                }
                else
                {
                    return false;
                }
            }
            else if ((ix=token.IndexOf('!')) > 1)
            {
                if (token.IndexOfAny(SheetNameInvalidChars, 0, token.IndexOf('!')) > 0)
                {
                    return false;
                }
                token = token.Substring(token.IndexOf('!') + 1);
            }
            return OfficeOpenXml.ExcelAddress.IsValidAddress(token);
        }
    }
}
