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
 * Mats Alm   		                Added		                2014-01-06
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Rounddown : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var number = ArgToDecimal(arguments, 0);
            var nDecimals = ArgToInt(arguments, 1);

            var nFactor = number < 0 ? -1 : 1;
            number *= nFactor;

            double result;
            if (nDecimals > 0)
            {
                result = RoundDownDecimalNumber(number, nDecimals);
            }
            else
            {
                result = (int)System.Math.Floor(number);
                result = result - (result % System.Math.Pow(10, (nDecimals*-1)));
            }
            return CreateResult(result * nFactor, DataType.Decimal);
        }

        private static double RoundDownDecimalNumber(double number, int nDecimals)
        {
            var integerPart = System.Math.Floor(number);
            var decimalPart = number - integerPart;
            decimalPart = System.Math.Pow(10d, nDecimals)*decimalPart;
            decimalPart = System.Math.Truncate(decimalPart)/System.Math.Pow(10d, nDecimals);
            var result = integerPart + decimalPart;
            return result;
        }
    }
}
