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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class SumIf : HiddenValuesHandlingFunction
    {
        private readonly ExpressionEvaluator _evaluator;

        public SumIf()
            : this(new ExpressionEvaluator())
        {

        }

        public SumIf(ExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            _evaluator = evaluator;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var args = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
            var criteria = arguments.ElementAt(1).ValueFirst != null ? ArgToString(arguments, 1) : null;
            var retVal = 0d;
            if (args == null)
            {
                var val = arguments.ElementAt(0).Value;
                if (criteria != null && _evaluator.Evaluate(val, criteria))
                {
                    var sumRange = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
                    retVal = arguments.Count() > 2
                        ? sumRange.First().ValueDouble
                        : ConvertUtil.GetValueDouble(val, true);
                }
            }
            else if (arguments.Count() > 2)
            {
                var sumRange = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
                retVal = CalculateWithSumRange(args, criteria, sumRange, context);
            }
            else
            {
                retVal = CalculateSingleRange(args, criteria, context);
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double CalculateWithSumRange(ExcelDataProvider.IRangeInfo range, string criteria, ExcelDataProvider.IRangeInfo sumRange, ParsingContext context)
        {
            var retVal = 0d;
            foreach (var cell in range)
            {
                if (criteria != null && _evaluator.Evaluate(cell.Value, criteria))
                {
                    var or = cell.Row - range.Address._fromRow;
                    var oc = cell.Column - range.Address._fromCol;
                    if (sumRange.Address._fromRow + or <= sumRange.Address._toRow &&
                       sumRange.Address._fromCol + oc <= sumRange.Address._toCol)
                    {
                        var v = sumRange.GetOffset(or, oc);
                        if (v is ExcelErrorValue)
                        {
                            throw (new ExcelErrorValueException((ExcelErrorValue)v));
                        }
                        retVal += ConvertUtil.GetValueDouble(v, true);
                    }
                }
            }
            return retVal;
        }

        private double CalculateSingleRange(ExcelDataProvider.IRangeInfo range, string expression, ParsingContext context)
        {
            var retVal = 0d;
            foreach (var candidate in range)
            {
                if (expression != null && IsNumeric(candidate.Value) && _evaluator.Evaluate(candidate.Value, expression) && IsNumeric(candidate.Value))
                {
                    if (candidate.IsExcelError)
                    {
                        throw (new ExcelErrorValueException((ExcelErrorValue)candidate.Value));
                    }
                    retVal += candidate.ValueDouble;
                }
            }
            return retVal;
        }
    }
}
