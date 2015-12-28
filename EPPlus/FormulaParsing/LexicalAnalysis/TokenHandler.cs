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
 * Mats Alm   		                Added       		        2015-12-28
 *******************************************************************************/
using OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class TokenHandler : ITokenIndexProvider
    {
        public TokenHandler(TokenizerContext context, ITokenFactory tokenFactory, ITokenSeparatorProvider tokenProvider)
        {
            _context = context;
            _tokenFactory = tokenFactory;
            _tokenProvider = tokenProvider;
        }

        private readonly TokenizerContext _context;
        private readonly ITokenSeparatorProvider _tokenProvider;
        private readonly ITokenFactory _tokenFactory;
        private int _tokenIndex = -1;

        public string Worksheet { get; set; }

        public bool HasMore()
        {
            return _tokenIndex < (_context.FormulaChars.Length - 1);
        }

        public void Next()
        {
            _tokenIndex++;
            Handle();
        }

        private void Handle()
        {
            var c = _context.FormulaChars[_tokenIndex];
            Token tokenSeparator;
            if (CharIsTokenSeparator(c, out tokenSeparator))
            {
                if (TokenSeparatorHandler.Handle(c, tokenSeparator, _context, this))
                {
                    return;
                }
                                              
                if (_context.CurrentTokenHasValue)
                {
                    if (Regex.IsMatch(_context.CurrentToken, "^\"*$"))
                    {
                        _context.AddToken(_tokenFactory.Create(_context.CurrentToken, TokenType.StringContent));
                    }
                    else
                    {
                        _context.AddToken(CreateToken(_context, Worksheet));
                    }


                    //If the a next token is an opening parantheses and the previous token is interpeted as an address or name, then the currenct token is a function
                    if (tokenSeparator.TokenType == TokenType.OpeningParenthesis && (_context.LastToken.TokenType == TokenType.ExcelAddress || _context.LastToken.TokenType == TokenType.NameValue))
                    {
                        _context.LastToken.TokenType = TokenType.Function;
                    }
                }
                if (tokenSeparator.Value == "-")
                {
                    if (TokenIsNegator(_context))
                    {
                        _context.AddToken(new Token("-", TokenType.Negator));
                        return;
                    }
                }
                _context.AddToken(tokenSeparator);
                _context.NewToken();
                return;
            }
            _context.AppendToCurrentToken(c);
        }

        private bool CharIsTokenSeparator(char c, out Token token)
        {
            var result = _tokenProvider.Tokens.ContainsKey(c.ToString());
            token = result ? token = _tokenProvider.Tokens[c.ToString()] : null;
            return result;
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

        int ITokenIndexProvider.Index
        {
            get { return _tokenIndex; }
        }


        void ITokenIndexProvider.MoveIndexPointerForward()
        {
            _tokenIndex++;
        }
    }
}
