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
 * Mats Alm   		                Added		                2015-04-06
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    public class Dget : ExcelFunction
    {
        private readonly RowMatcher _rowMatcher;

        public Dget()
            : this(new RowMatcher())
        {
            
        }

        public Dget(RowMatcher rowMatcher)
        {
            _rowMatcher = rowMatcher;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var dbAddress = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;
            var field = ArgToString(arguments, 1).ToLower(CultureInfo.InvariantCulture);
            var criteriaRange = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;

            var db = new ExcelDatabase(context.ExcelDataProvider, dbAddress);
            var criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, criteriaRange);

            var nHits = 0;
            object retVal = null;
            while (db.HasMoreRows)
            {
                var dataRow = db.Read();
                if (_rowMatcher.IsMatch(dataRow, criteria.Items))
                {
                    if(++nHits > 1) return CreateResult(ExcelErrorValue.Values.Num, DataType.ExcelError);
                    retVal = dataRow[field];
                }
            }
            return new CompileResultFactory().Create(retVal);
        }
    }
}
