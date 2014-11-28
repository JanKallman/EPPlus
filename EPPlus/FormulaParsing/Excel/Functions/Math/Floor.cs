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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Floor : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var number = ArgToDecimal(arguments, 0);
            var significance = ArgToDecimal(arguments, 1);
            ValidateNumberAndSign(number, significance);
            if (significance < 1 && significance > 0)
            {
                var floor = System.Math.Floor(number);
                var rest = number - floor;
                var nSign = (int)(rest / significance);
                return CreateResult(floor + (nSign * significance), DataType.Decimal);
            }
            else if (significance == 1)
            {
                return CreateResult(System.Math.Floor(number), DataType.Decimal);
            }
            else
            {
                double result;
                if (number > 1)
                {
                    result = number - (number % significance) + significance;
                }
                else
                {
                    result = number - (number % significance);
                }
                return CreateResult(result, DataType.Decimal);
            }
        }

        private void ValidateNumberAndSign(double number, double sign)
        {
            if (number > 0d && sign < 0)
            {
                var values = string.Format("num: {0}, sign: {1}", number, sign);
                throw new InvalidOperationException("Floor cannot handle a negative significance when the number is positive" + values);
            }
        }
    }
}
