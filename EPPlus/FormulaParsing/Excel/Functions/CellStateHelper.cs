/* Copyright (C) 2011  Jan Källman
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
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
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

        internal static bool ShouldIgnore(bool ignoreHiddenValues, FunctionArgument arg, ParsingContext context)
        {
            return (ignoreHiddenValues && arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
        }
    }
}
