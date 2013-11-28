using System;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    internal static class CellStateHelper
    {
        private static bool IsSubTotal(ExcelDataProvider.ICellInfo c)
        {
            var tokens = c.Tokens;
            if (tokens == null) return false;
            return c.Tokens.Any(token => 
                token.TokenType == LexicalAnalysis.TokenType.Function 
                && token.Value.Equals("SUBTOTAL", StringComparison.InvariantCultureIgnoreCase)
                );
        }

        internal static bool ShouldIgnore(bool ignoreHiddenValues, ExcelDataProvider.ICellInfo c, ParsingContext context)
        {
            return (ignoreHiddenValues && c.IsHiddenRow) || (context.Scopes.Current.IsSubtotal && IsSubTotal(c));
        }
    }
}
