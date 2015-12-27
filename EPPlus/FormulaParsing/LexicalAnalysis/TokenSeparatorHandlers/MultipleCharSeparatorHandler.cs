using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    public class MultipleCharSeparatorHandler : SeparatorHandler
    {
        ITokenSeparatorProvider _tokenSeparatorProvider;

        public MultipleCharSeparatorHandler()
            : this(new TokenSeparatorProvider())
        {

        }
        public MultipleCharSeparatorHandler(ITokenSeparatorProvider tokenSeparatorProvider)
        {
            _tokenSeparatorProvider = tokenSeparatorProvider;
        }
        public override bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            // two operators in sequence could be "<=" or ">="
            if (IsPartOfMultipleCharSeparator(context, c))
            {
                var sOp = context.LastToken.Value + c.ToString(CultureInfo.InvariantCulture);
                var op = _tokenSeparatorProvider.Tokens[sOp];
                context.ReplaceLastToken(op);
                context.NewToken();
                return true;
            }
            return false;
        }

        private bool IsPartOfMultipleCharSeparator(TokenizerContext context, char c)
        {
            var lastToken = context.LastToken != null ? context.LastToken.Value : string.Empty;
            return _tokenSeparatorProvider.IsOperator(lastToken)
                && _tokenSeparatorProvider.IsPossibleLastPartOfMultipleCharOperator(c.ToString(CultureInfo.InvariantCulture))
                && !context.CurrentTokenHasValue;
        }
    }
}
