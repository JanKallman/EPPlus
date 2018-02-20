/* Copyright (C) 2017  Abedilah Hassan
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
 * Abedilah Hassan   		        Added		               2018/02/20
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using Microsoft.VisualBasic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Financial
{
    class Rate : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var Nper = ArgToDecimal(arguments, 0);
            var pmt = ArgToDecimal(arguments, 1);
            var Pv = ArgToDecimal(arguments, 2);
            double Fv = 0;
            if (arguments.Count() > 3)
                Fv = ArgToDecimal(arguments, 3);
            DueDate dueDate = 0;
            if (arguments.Count() > 4)
            {
                var temp = ArgToDecimal(arguments, 0);
                if (temp == 0)
                {
                    dueDate = DueDate.BegOfPeriod;
                }
                else
                {
                    dueDate = DueDate.EndOfPeriod;
                }
            }
            double guess = 0;
            if (arguments.Count() > 5)
            {
                guess = ArgToDecimal(arguments, 5);
            }

            return CreateResult(Microsoft.VisualBasic.Financial.Rate(Nper, pmt, Pv, Fv, dueDate, guess), DataType.Decimal);
        }
    }
}
