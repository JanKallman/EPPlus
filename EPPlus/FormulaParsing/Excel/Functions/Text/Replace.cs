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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Replace : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 4);
            var oldText = ArgToString(arguments, 0);
            var startPos = ArgToInt(arguments, 1);
            var nCharsToReplace = ArgToInt(arguments, 2);
            var newText = ArgToString(arguments, 3);
            var firstPart = GetFirstPart(oldText, startPos);
            var lastPart = GetLastPart(oldText, startPos, nCharsToReplace);
            var result = string.Concat(firstPart, newText, lastPart);
            return CreateResult(result, DataType.String);
        }

        private string GetFirstPart(string text, int startPos)
        {
            return text.Substring(0, startPos - 1);
        }

        private string GetLastPart(string text, int startPos, int nCharactersToReplace)
        {
            int startIx = startPos -1;
            startIx += nCharactersToReplace;
            return text.Substring(startIx, text.Length - startIx);
        }
    }
}
