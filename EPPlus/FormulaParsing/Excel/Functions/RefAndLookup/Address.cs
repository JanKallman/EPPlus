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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class Address : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var row = ArgToInt(arguments, 0);
            var col = ArgToInt(arguments, 1);
            ThrowExcelFunctionExceptionIf(() => row < 0 && col < 0, ExcelErrorCodes.Value);
            var referenceType = ExcelReferenceType.AbsoluteRowAndColumn;
            var worksheetSpec = string.Empty;
            if (arguments.Count() > 2)
            {
                var arg3 = ArgToInt(arguments, 2);
                ThrowExcelFunctionExceptionIf(() => arg3 < 1 || arg3 > 4, ExcelErrorCodes.Value);
                referenceType = (ExcelReferenceType)ArgToInt(arguments, 2);
            }
            if (arguments.Count() > 3)
            {
                var fourthArg = arguments.ElementAt(3).Value;
                if(fourthArg.GetType().Equals(typeof(bool)) && !(bool)fourthArg)
                {
                    throw new InvalidOperationException("Excelformulaparser does not support the R1C1 format!");
                }
                if (fourthArg.GetType().Equals(typeof(string)))
                {
                    worksheetSpec = fourthArg.ToString() + "!";
                }
            }
            var translator = new IndexToAddressTranslator(context.ExcelDataProvider, referenceType);
            return CreateResult(worksheetSpec + translator.ToAddress(col, row), DataType.ExcelAddress);
        }
    }
}
