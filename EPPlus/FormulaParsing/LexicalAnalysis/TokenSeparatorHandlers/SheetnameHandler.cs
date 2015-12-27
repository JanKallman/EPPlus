using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    public class SheetnameHandler : SeparatorHandler
    {
        public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            if (context.IsInSheetName)
            {
                if (IsDoubleQuote(tokenSeparator, tokenIndexProvider.Index, context))
                {
                    tokenIndexProvider.MoveIndexPointerForward();
                    context.AppendToCurrentToken(c);
                    return true;
                }
                if (tokenSeparator.TokenType != TokenType.WorksheetName)
                {
                    context.AppendToCurrentToken(c);
                    return true;
                }
            }

            if (tokenSeparator.TokenType == TokenType.WorksheetName)
            {
                if (context.LastToken != null && context.LastToken.TokenType == TokenType.WorksheetName)
                {
                    context.AddToken(!context.CurrentTokenHasValue
                        ? new Token(string.Empty, TokenType.WorksheetNameContent)
                        : new Token(context.CurrentToken, TokenType.WorksheetNameContent));
                }
                context.AddToken(new Token("'", TokenType.WorksheetName));
                context.ToggleIsInSheetName();
                context.NewToken();
                return true;
            }
            return false;
        }
    }
}
