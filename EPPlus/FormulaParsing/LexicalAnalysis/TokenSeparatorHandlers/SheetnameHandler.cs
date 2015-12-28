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
