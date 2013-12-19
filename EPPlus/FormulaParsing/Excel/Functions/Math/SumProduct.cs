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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class SumProduct : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            double result = 0d;
            List<List<double>> results = new List<List<double>>();
            foreach (var arg in arguments)
            {
                results.Add(new List<double>());
                var currentResult = results.Last();
                if (arg.Value is IEnumerable<FunctionArgument>)
                {
                    foreach (var val in (IEnumerable<FunctionArgument>)arg.Value)
                    {
                        AddValue(val.Value, currentResult);
                    }
                }
                else if (arg.IsExcelRange)
                {
                    foreach (var val in arg.ValueAsRangeInfo)
                    {
                        AddValue(val.Value, currentResult);
                    }
                }
            }
            // Validate that all supplied lists have the same length
            var arrayLength = results.First().Count;
            foreach (var list in results)
            {
                if (list.Count != arrayLength)
                {
                    throw new ExcelErrorValueException(ExcelErrorValue.Create(eErrorType.Value));
                    //throw new ExcelFunctionException("All supplied arrays must have the same length", ExcelErrorCodes.Value);
                }
            }
            for (var rowIndex = 0; rowIndex < arrayLength; rowIndex++)
            {
                double rowResult = 1;
                for (var colIndex = 0; colIndex < results.Count; colIndex++)
                {
                    rowResult *= results[colIndex][rowIndex];
                }
                result += rowResult;
            }
            return CreateResult(result, DataType.Decimal);
        }

        private void AddValue(object convertVal, List<double> currentResult)
        {
            if (IsNumeric(convertVal))
            {
                currentResult.Add(Convert.ToDouble(convertVal));
            }
            else
            {
                currentResult.Add(0d);
            }
        }
    }
}
