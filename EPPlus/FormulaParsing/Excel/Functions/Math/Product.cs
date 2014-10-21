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
    public class Product : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var result = 0d;
            var index = 0;
            while (AreEqual(result, 0d) && index < arguments.Count())
            {
                result = CalculateFirstItem(arguments, index++, context);
            }
            result = CalculateCollection(arguments.Skip(index), result, (arg, current) =>
            {
                if (ShouldIgnore(arg)) return current;
                if (arg.ValueIsExcelError)
                {
                    ThrowExcelErrorValueException(arg.ValueAsExcelErrorValue.Type);
                }
                if (arg.IsExcelRange)
                {
                    foreach (var cell in arg.ValueAsRangeInfo)
                    {
                        if(ShouldIgnore(cell, context)) return current;
                        current *= cell.ValueDouble;
                    }
                    return current;
                }
                var obj = arg.Value;
                if (obj != null && IsNumeric(obj))
                {
                    var val = Convert.ToDouble(obj);
                    current *= val;
                }
                return current;
            });
            return CreateResult(result, DataType.Decimal);
        }

        private double CalculateFirstItem(IEnumerable<FunctionArgument> arguments, int index, ParsingContext context)
        {
            var element = arguments.ElementAt(index);
            var argList = new List<FunctionArgument> { element };
            var valueList = ArgsToDoubleEnumerable(false, false, argList, context);
            var result = 0d;
            foreach (var value in valueList)
            {
                if (result == 0d && value > 0d)
                {
                    result = value;
                }
                else
                {
                    result *= value;
                }
            }
            return result;
        }
    }
}
