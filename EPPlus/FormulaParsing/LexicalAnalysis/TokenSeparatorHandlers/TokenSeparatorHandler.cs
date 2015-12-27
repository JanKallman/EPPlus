using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis.TokenSeparatorHandlers
{
    public static class TokenSeparatorHandler
    {
        private static SeparatorHandler[] _handlers = new SeparatorHandler[]
        { 
            new StringHandler(),
            new BracketHandler(),
            new SheetnameHandler(),
            new MultipleCharSeparatorHandler()
        };

        public static bool Handle(char c, Token tokenSeparator, TokenizerContext context, ITokenIndexProvider tokenIndexProvider)
        {
            foreach(var handler in _handlers)
            {
                if(handler.Handle(c, tokenSeparator, context, tokenIndexProvider))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
