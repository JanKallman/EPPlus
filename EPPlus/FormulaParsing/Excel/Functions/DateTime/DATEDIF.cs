/* Copyright (C) 2017  
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
 * Abedilah Hassan   		        Added		               2018/02/26
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using Microsoft.VisualBasic;
namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
    class DATEDIF : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var startDate = ArgToDecimal(arguments, 0);
            var endDate = ArgToDecimal(arguments, 1);
            var interval = ArgToString(arguments, 2).ToString();
            Microsoft.VisualBasic.DateInterval dateInterval= Microsoft.VisualBasic.DateInterval.Year;
            if(interval.ToLower().Equals("y"))
            {
                dateInterval = Microsoft.VisualBasic.DateInterval.Year;
            }
            if (interval.ToLower().Equals("m"))
            {
                dateInterval = Microsoft.VisualBasic.DateInterval.Month;
            }
            if (interval.ToLower().Equals("d"))
            {
                dateInterval = Microsoft.VisualBasic.DateInterval.Day;
            }
            if (interval.ToLower().Equals("h"))
            {
                dateInterval = Microsoft.VisualBasic.DateInterval.Hour;
            }
            if (interval.ToLower().Equals("n"))
            {
                dateInterval = Microsoft.VisualBasic.DateInterval.Minute;
            }
            if (interval.ToLower().Equals("s"))
            {
                dateInterval = Microsoft.VisualBasic.DateInterval.Second;
            }
           return CreateResult(Microsoft.VisualBasic.DateAndTime.DateDiff(dateInterval,
               System.DateTime.FromOADate(startDate), System.DateTime.FromOADate(endDate)), DataType.Integer);
        }
    }
}
