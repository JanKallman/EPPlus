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
 * Mats Alm   		                Added		                2015-01-11
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class Offset : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 3);
            var startRange = ArgToAddress(functionArguments, 0, context);
            var rowOffset = ArgToInt(functionArguments, 1);
            var colOffset = ArgToInt(functionArguments, 2);
            int width = 0, height = 0;
            if (functionArguments.Length > 3)
            {
                height = ArgToInt(functionArguments, 3);
                if (height == 0) return new CompileResult(eErrorType.Ref);
            }
            if (functionArguments.Length > 4)
            {
                width = ArgToInt(functionArguments, 4);
                if (width == 0) return new CompileResult(eErrorType.Ref);
            }
            var ws = context.Scopes.Current.Address.Worksheet;            
            var r =context.ExcelDataProvider.GetRange(ws,startRange);
            var adr = r.Address;

            var fromRow = adr._fromRow + rowOffset;
            var fromCol = adr._fromCol + colOffset;
            var toRow = (height != 0 ? adr._fromRow + height - 1 : adr._toRow) + rowOffset;
            var toCol = (width != 0 ? adr._fromCol + width - 1 : adr._toCol) + colOffset;

            var newRange = context.ExcelDataProvider.GetRange(adr.WorkSheet, fromRow, fromCol, toRow, toCol);
            
            return CreateResult(newRange, DataType.Enumerable);
        }
    }
}
