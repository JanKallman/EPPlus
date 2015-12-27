using OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
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
