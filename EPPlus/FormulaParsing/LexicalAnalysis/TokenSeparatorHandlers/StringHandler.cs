using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    public class StringHandler : SeparatorHandler
    {
        public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            if(context.IsInString)
            { 
                if (IsDoubleQuote(tokenSeparator, tokenIndexProvider.Index, context))
                {
                    tokenIndexProvider.MoveIndexPointerForward();
                    context.AppendToCurrentToken(c);
                    return true;
                }
                if (tokenSeparator.TokenType != TokenType.String)
                {
                    context.AppendToCurrentToken(c);
                    return true;
                }
            }

            if (tokenSeparator.TokenType == TokenType.String)
            {
                if (context.LastToken != null && context.LastToken.TokenType == TokenType.OpeningEnumerable)
                {
                    context.AppendToCurrentToken(c);
                    context.ToggleIsInString();
                    return true;
                }
                if (context.LastToken != null && context.LastToken.TokenType == TokenType.String)
                {
                    context.AddToken(!context.CurrentTokenHasValue
                        ? new Token(string.Empty, TokenType.StringContent)
                        : new Token(context.CurrentToken, TokenType.StringContent));
                }
                context.AddToken(new Token("\"", TokenType.String));
                context.ToggleIsInString();
                context.NewToken();
                return true;
            }
            return false;
        }
    }
}
