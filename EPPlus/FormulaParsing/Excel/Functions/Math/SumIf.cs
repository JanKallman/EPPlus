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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using Util=OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class SumIf : HiddenValuesHandlingFunction
    {
        private readonly NumericExpressionEvaluator _evaluator;

        public SumIf()
            : this(new NumericExpressionEvaluator())
        {

        }

        public SumIf(NumericExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            _evaluator = evaluator;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var args = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo; //IEnumerable<FunctionArgument>;
            var criteria = arguments.ElementAt(1).Value;
            ThrowExcelErrorValueExceptionIf(() => criteria == null || criteria.ToString().Length > 255, eErrorType.Value);
            var retVal = 0d;
            if (arguments.Count() > 2)
            {
                var sumRange = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;//IEnumerable<FunctionArgument>;
                retVal = CalculateWithSumRange(args, criteria.ToString(), sumRange, context);
            }
            else
            {
                retVal = CalculateSingleRange(args, criteria.ToString(), context);
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double CalculateWithSumRange(ExcelDataProvider.IRangeInfo range, string criteria, ExcelDataProvider.IRangeInfo sumRange, ParsingContext context)
        {
            var retVal = 0d;
            foreach(var cell in range)
            {
                if (_evaluator.Evaluate(cell.Value, criteria))
                {
                    var or = cell.Row-range.Address._fromRow;
                    var oc = cell.Column - range.Address._fromCol;
                    if(sumRange.Address._fromRow+or <= sumRange.Address._toRow && 
                       sumRange.Address._fromCol+oc <= sumRange.Address._toCol)
                    {
                        var v = sumRange.GetOffset(or, oc);
                        if (v is ExcelErrorValue)
                        {
                            throw (new ExcelErrorValueException((ExcelErrorValue)v));
                        }
                        retVal += Util.ConvertUtil.GetValueDouble(v, true);
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
                if (_evaluator.Evaluate(candidate.Value, expression))
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

        //private double Calculate(FunctionArgument arg, string expression)
        //{
        //    var retVal = 0d;
        //    if (ShouldIgnore(arg) || !_evaluator.Evaluate(arg.Value, expression))
        //    {
        //        return retVal;
        //    }
        //    if (arg.Value is double || arg.Value is int)
        //    {
        //        retVal += Convert.ToDouble(arg.Value);
        //    }
        //    else if (arg.Value is System.DateTime)
        //    {
        //        retVal += Convert.ToDateTime(arg.Value).ToOADate();
        //    }
        //    else if (arg.Value is IEnumerable<FunctionArgument>)
        //    {
        //        foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
        //        {
        //            retVal += Calculate(item, expression);
        //        }
        //    }
        //    return retVal;
        //}
    }
}
