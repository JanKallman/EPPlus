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
using OfficeOpenXml.FormulaParsing.Exceptions;
using MathObj = System.Math;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class Stdev : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var values = ArgsToDoubleEnumerable(arguments, context, false);
            return CreateResult(StandardDeviation(values), DataType.Decimal);
        }

        private static double StandardDeviation(IEnumerable<double> values)
        {
            double ret = 0;
            if (values.Count() > 0)
            {
                var nValues = values.Count();
                if(nValues == 1) throw new ExcelErrorValueException(eErrorType.Div0);
                //Compute the Average       
                double avg = values.Average();
                //Perform the Sum of (value-avg)_2_2       
                double sum = values.Sum(d => MathObj.Pow(d - avg, 2));
                //Put it all together       
                ret = MathObj.Sqrt((sum) / (values.Count() - 1));
            }
            return ret;
        } 

    }
}
