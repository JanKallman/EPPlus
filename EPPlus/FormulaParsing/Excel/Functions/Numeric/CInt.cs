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
using System.Globalization;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric
{
    public class CInt : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var obj = arguments.ElementAt(0).ValueFirst;
            return CreateResult(ToInteger(obj), DataType.Integer);
        }

        private object ToInteger(object obj)
        {
            var type = obj.GetType();
            if (type == typeof(double))
            {
                return (int)System.Math.Floor((double)obj);
            }
            if (type == typeof(decimal))
            {
                return (int)System.Math.Floor((decimal)obj);
            }
            double result;
            if(double.TryParse(HandleDecimalSeparator(obj), out result))
            {
                return (int)System.Math.Floor(result);
            }
            throw new ArgumentException("Could not cast supplied argument to integer");
        }

        private string HandleDecimalSeparator(object obj)
        {
            var separator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            string val = obj != null ? obj.ToString() : string.Empty;
            if (separator == ",")
            {
                val = val.Replace(".", ",");
            }
            return val;
        }
    }
}
