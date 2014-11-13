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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class AverageA : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1, eErrorType.Div0);
            double nValues = 0d, result = 0d;
            foreach (var arg in arguments)
            {
                Calculate(arg, context, ref result, ref nValues);
            }
            if (nValues == 0d) ThrowExcelErrorValueException(eErrorType.Div0);
            return CreateResult(result / nValues, DataType.Decimal);
        }

        private void Calculate(FunctionArgument arg, ParsingContext context, ref double retVal, ref double nValues, bool isInArray = false)
        {
            if (ShouldIgnore(arg))
            {
                return;
            }
            if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    Calculate(item, context, ref retVal, ref nValues, true);
                }
            }
            else if (arg.IsExcelRange)
            {
                foreach (var c in arg.ValueAsRangeInfo)
                {
                    if (ShouldIgnore(c, context)) continue;
                    CheckForAndHandleExcelError(c);
                    if (IsNumeric(c.Value))
                    {
                        nValues++;
                        retVal += c.ValueDouble;
                    }
                    else if (c.Value is bool)
                    {
                        nValues++;
                        retVal += (bool) c.Value ? 1 : 0;
                    }
                    else if (c.Value is string)
                    {
                        nValues++;
                    }
                }
            }
            else
            {
                var numericValue = GetNumericValue(arg.Value, isInArray);
                if (numericValue.HasValue)
                {
                    nValues++;
                    retVal += numericValue.Value;
                }
                else if((arg.Value is string) && !ConvertUtil.IsNumericString(arg.Value))
                {
                    if (isInArray)
                    {
                        nValues++;
                    }
                    else
                    {
                        ThrowExcelErrorValueException(eErrorType.Value);   
                    }
                }
            }
            CheckForAndHandleExcelError(arg);
        }

        private double? GetNumericValue(object obj, bool isInArray)
        {
            if (IsNumeric(obj))
            {
                return ConvertUtil.GetValueDouble(obj);
            }
            else if (obj is bool)
            {
                if (isInArray) return default(double?);
                return ConvertUtil.GetValueDouble(obj);
            }
            else if (ConvertUtil.IsNumericString(obj))
            {
                return double.Parse(obj.ToString(), CultureInfo.InvariantCulture);
            }
            return default(double?);
        }
    }
}
