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
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    public class DcountA : DatabaseFunction
    {

        public DcountA()
            : this(new RowMatcher())
        {

        }

        public DcountA(RowMatcher rowMatcher)
            : base(rowMatcher)
        {

        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var dbAddress = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;
            string field = null;
            string criteriaRange = null;
            if (arguments.Count() == 2)
            {
                criteriaRange = arguments.ElementAt(1).ValueAsRangeInfo.Address.Address;
            }
            else
            {
                field = ArgToString(arguments, 1).ToLower(CultureInfo.InvariantCulture);
                criteriaRange = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;
            }
            var db = new ExcelDatabase(context.ExcelDataProvider, dbAddress);
            var criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, criteriaRange);

            var nHits = 0;
            while (db.HasMoreRows)
            {
                var dataRow = db.Read();
                if (RowMatcher.IsMatch(dataRow, criteria.Items))
                {
                    // if a fieldname is supplied, count only this row if the value
                    // of the supplied field is not blank.
                    if (!string.IsNullOrEmpty(field))
                    {
                        var candidate = dataRow[field];
                        if (ShouldCount(candidate))
                        {
                            nHits++;
                        }
                    }
                    else
                    {
                        // no fieldname was supplied, always count matching row.
                        nHits++;
                    }
                }
            }
            return CreateResult(nHits, DataType.Integer);
        }

        private bool ShouldCount(object value)
        {
            if (value == null) return false;
            return (!string.IsNullOrEmpty(value.ToString()));
        }
    }
}
