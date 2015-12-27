using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    public abstract class SeparatorHandler
    {
        protected bool IsDoubleQuote(Token tokenSeparator, int formulaCharIndex, TokenizerContext context)
        {
            return tokenSeparator.TokenType == TokenType.String && formulaCharIndex + 1 < context.FormulaChars.Length && context.FormulaChars[formulaCharIndex + 1] == '\"';
        }

        public abstract bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider);

    }
}
