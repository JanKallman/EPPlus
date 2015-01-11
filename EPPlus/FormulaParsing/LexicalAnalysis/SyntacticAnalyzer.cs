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
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class SyntacticAnalyzer : ISyntacticAnalyzer
    {
        private class AnalyzingContext
        {
            public int NumberOfOpenedParentheses { get; set; }
            public int NumberOfClosedParentheses { get; set; }
            public int OpenedStrings { get; set; }
            public int ClosedStrings { get; set; }
            public bool IsInString { get; set; }
        }
        public void Analyze(IEnumerable<Token> tokens)
        {
            var context = new AnalyzingContext();
            foreach (var token in tokens)
            {
                if (token.TokenType == TokenType.Unrecognized)
                {
                    throw new UnrecognizedTokenException(token);
                }
                EnsureParenthesesAreWellFormed(token, context);
                EnsureStringsAreWellFormed(token, context);
            }
            Validate(context);
        }

        private static void Validate(AnalyzingContext context)
        {
            if (context.NumberOfOpenedParentheses != context.NumberOfClosedParentheses)
            {
                throw new FormatException("Number of opened and closed parentheses does not match");
            }
            if (context.OpenedStrings != context.ClosedStrings)
            {
                throw new FormatException("Unterminated string");
            }
        }

        private void EnsureParenthesesAreWellFormed(Token token, AnalyzingContext context)
        {
            if (token.TokenType == TokenType.OpeningParenthesis)
            {
                context.NumberOfOpenedParentheses++;
            }
            else if (token.TokenType == TokenType.ClosingParenthesis)
            {
                context.NumberOfClosedParentheses++;
            }
        }

        private void EnsureStringsAreWellFormed(Token token, AnalyzingContext context)
        {
            if (!context.IsInString && token.TokenType == TokenType.String)
            {
                context.IsInString = true;
                context.OpenedStrings++;
            }
            else if (context.IsInString && token.TokenType == TokenType.String)
            {
                context.IsInString = false;
                context.ClosedStrings++;
            }
        }
    }
}
