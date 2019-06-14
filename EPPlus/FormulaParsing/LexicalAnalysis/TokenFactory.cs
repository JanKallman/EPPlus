﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class TokenFactory : ITokenFactory
    {
        public TokenFactory(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider, bool r1c1=false)
            : this(new TokenSeparatorProvider(), nameValueProvider, functionRepository, r1c1)
        {

        }

        public TokenFactory(ITokenSeparatorProvider tokenSeparatorProvider, INameValueProvider nameValueProvider, IFunctionNameProvider functionNameProvider, bool r1c1)
        {
            _tokenSeparatorProvider = tokenSeparatorProvider;
            _functionNameProvider = functionNameProvider;
            _nameValueProvider = nameValueProvider;
            _r1c1 = r1c1;
        }

        private readonly ITokenSeparatorProvider _tokenSeparatorProvider;
        private readonly IFunctionNameProvider _functionNameProvider;
        private readonly INameValueProvider _nameValueProvider;
        private bool _r1c1;
        public Token Create(IEnumerable<Token> tokens, string token)
        {
            return Create(tokens, token, null);
        }
        public Token Create(IEnumerable<Token> tokens, string token, string worksheet)
        {
            Token tokenSeparator = null;
            if (_tokenSeparatorProvider.Tokens.TryGetValue(token, out tokenSeparator))
            {
                return tokenSeparator;
            }
            var tokenList = (IList<Token>)tokens;
            //Address with worksheet-string before  /JK
            if (token.StartsWith("!") && tokenList[tokenList.Count-1].TokenType == TokenType.String)
            {
                string addr = "";
                var i = tokenList.Count - 2;
                if (i > 0)
                {
                    if (tokenList[i].TokenType == TokenType.StringContent)
                    {
                        addr = "'" + tokenList[i].Value.Replace("'", "''") + "'";
                    }
                    else
                    {
                        throw(new ArgumentException(string.Format("Invalid formula token sequence near {0}",token)));
                    }
                    //Remove the string tokens and content
                    tokenList.RemoveAt(tokenList.Count - 1);
                    tokenList.RemoveAt(tokenList.Count - 1);
                    tokenList.RemoveAt(tokenList.Count - 1);

                    return new Token(addr + token, TokenType.ExcelAddress);
                }
                else
                {
                    throw(new ArgumentException(string.Format("Invalid formula token sequence near {0}",token)));
                }
                
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
            if (token.ToUpper(CultureInfo.InvariantCulture).Contains("#REF!"))
            {
                return new Token(token, TokenType.InvalidReference);
            }
            if (token.ToUpper(CultureInfo.InvariantCulture) == "#NUM!")
            {
                return new Token(token, TokenType.NumericError);
            }
            if (token.ToUpper(CultureInfo.InvariantCulture) == "#VALUE!")
            {
                return new Token(token, TokenType.ValueDataTypeError);
            }
            if (token.ToUpper(CultureInfo.InvariantCulture) == "#NULL!")
            {
                return new Token(token, TokenType.Null);
            }
            if (_nameValueProvider != null && _nameValueProvider.IsNamedValue(token, worksheet))
            {
                return new Token(token, TokenType.NameValue);
            }
            if (_functionNameProvider.IsFunctionName(token))
            {
                return new Token(token, TokenType.Function);
            }
            if (tokenList.Count > 0 && tokenList[tokenList.Count - 1].TokenType == TokenType.OpeningEnumerable)
            {
                return new Token(token, TokenType.Enumerable);
            }
            var at = OfficeOpenXml.ExcelAddressBase.IsValid(token, out string normalizedToken, _r1c1);
            if (at==ExcelAddressBase.AddressType.InternalAddress)
            {
                return new Token(normalizedToken, TokenType.ExcelAddress);
            } 
            else if (at == ExcelAddressBase.AddressType.R1C1)
            {
                return new Token(token.ToUpper(CultureInfo.InvariantCulture), TokenType.ExcelAddressR1C1);
            }
            return new Token(token, TokenType.Unrecognized);

        }

        public Token Create(string token, TokenType explicitTokenType)
        {
            return new Token(token, explicitTokenType);
        }
    }
}
