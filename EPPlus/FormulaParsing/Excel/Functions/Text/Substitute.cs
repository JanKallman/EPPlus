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
    public class Substitute : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            int argCount;
            ValidateArguments(arguments, 3, out argCount);
            var text = ArgToString(arguments, 0);
            var find = ArgToString(arguments, 1);
            var replaceWith = ArgToString(arguments, 2);
            string result;
            if (argCount > 3)
            {
                var instanceNum = ArgToInt(arguments, 3);
                result = ReplaceFirst(text, find, replaceWith, instanceNum);
            }
            else
            {
                result = text.Replace(find, replaceWith);
            }
            return CreateResult(result, DataType.String);
        }

        /// <summary>
        /// Replaces only the Nth instance of substring.
        /// </summary>
        /// <param name="text">String to modify</param>
        /// <param name="search">Substring to look for</param>
        /// <param name="replace">Replacement for the matched substring</param>
        /// <param name="instanceNumber">One-based index of match to replace</param>
        /// <returns>Modified copy of parameter text where only the specified substring instance has been replaced</returns>
        private static string ReplaceFirst(string text, string search, string replace, int instanceNumber)
        {
            int pos = -1;
            for (int i=0; i<instanceNumber; i++)
            {
                pos = text.IndexOf(search, pos+1);
            }
            if (pos < 0)
            {
                return text;
            }
            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }
    }
}
