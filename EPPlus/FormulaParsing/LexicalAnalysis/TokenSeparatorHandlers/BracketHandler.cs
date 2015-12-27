using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    public class BracketHandler : SeparatorHandler
    {
        public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            if (tokenSeparator.TokenType == TokenType.OpeningBracket)
            {
                context.AppendToCurrentToken(c);
                context.BracketCount++;
                return true;
            }
            if (tokenSeparator.TokenType == TokenType.ClosingBracket)
            {
                context.AppendToCurrentToken(c);
                context.BracketCount--;
                return true;
            }
            if (context.BracketCount > 0)
            {
                context.AppendToCurrentToken(c);
                return true;
            }
            return false;
        }
    }
}
