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
 * Mats Alm   		                Added		                2018-07-17
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Finance
{
    public class Pmt : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var rate = ArgToDecimal(arguments, 0);
            var nPer = ArgToInt(arguments, 1);
            var presentValue = ArgToDecimal(arguments, 2);
            var payEndOfPeriod = false;
            var futureValue = 0d;
            if (arguments.Count() > 3) futureValue = ArgToDecimal(arguments, 3);
            if (arguments.Count() > 4) payEndOfPeriod = ArgToBool(arguments, 4);

            var result = (futureValue + presentValue * System.Math.Pow(rate + 1, nPer)) * rate
                      /
                   ((payEndOfPeriod ? rate + 1 : 1) * (1 - System.Math.Pow(rate + 1, nPer)));
       

            return CreateResult(result, DataType.Decimal);
        }

        private static double GetInterest(double rate, double remainingAmount)
        {
            return remainingAmount * rate;
        }
    }
}
