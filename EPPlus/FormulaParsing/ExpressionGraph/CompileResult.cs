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
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class CompileResult
    {
        private static CompileResult _empty = new CompileResult(null, DataType.Empty);
        public static CompileResult Empty
        {
            get { return _empty; }
        }

        public CompileResult(object result, DataType dataType)
        {
            Result = result;
            DataType = dataType;
        }
        public object Result
        {
            get;
            private set;
        }
        public double ResultNumeric
        {
            get
            {
                if (IsNumeric)
                {
                    return Result == null ? 0 :  Convert.ToDouble(Result);
                }
                else if(Result is DateTime)
                {
                    return ((DateTime)Result).ToOADate();
                }
                else if(Result is TimeSpan)
                {
                    return new DateTime(((TimeSpan)Result).Ticks).ToOADate();
                }
                else if (IsNumericString)
                {
                    return double.Parse(Result.ToString(), CultureInfo.InvariantCulture);
                }
                else
                {
                    return 0;
                }
            }
        }

        public DataType DataType
        {
            get;
            private set;
        }

        public bool IsNumeric
        {
            get { return DataType == DataType.Decimal || DataType == DataType.Integer || DataType == DataType.Empty; }
        }

        public bool IsNumericString
        {
            get
            {
                if (DataType == DataType.String && Result != null)
                {
                    return Regex.IsMatch(Result.ToString(), @"^[\d]+(\,[\d])?");
                }
                return false;
            }
        }

        public bool IsResultOfSubtotal { get; set; }
    }
}
