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
 * Mats Alm   		                Added		                2015-01-15
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;


namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public abstract class MultipleRangeCriteriasFunction : ExcelFunction
    {

        private readonly NumericExpressionEvaluator _numericExpressionEvaluator;
        private readonly WildCardValueMatcher _wildCardValueMatcher;

        protected MultipleRangeCriteriasFunction()
            :this(new NumericExpressionEvaluator(), new WildCardValueMatcher())
        {
            
        }

        protected MultipleRangeCriteriasFunction(NumericExpressionEvaluator evaluator, WildCardValueMatcher wildCardValueMatcher)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            Require.That(wildCardValueMatcher).Named("wildCardValueMatcher").IsNotNull();
            _numericExpressionEvaluator = evaluator;
            _wildCardValueMatcher = wildCardValueMatcher;
        }

        protected bool Evaluate(object obj, object expression)
        {
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue && expression is string)
            {
                return _numericExpressionEvaluator.Evaluate(candidate.Value, expression.ToString());
            }
            if (obj == null) return false;
            return _wildCardValueMatcher.IsMatch(expression, obj.ToString()) == 0;
        }

        protected List<int> GetMatchIndexes(ExcelDataProvider.IRangeInfo rangeInfo, object searched)
        {
            var result = new List<int>();
            var internalIndex = 0;
            for (var row = rangeInfo.Address._fromRow; row <= rangeInfo.Address._toRow; row++)
            {
                for (var col = rangeInfo.Address._fromCol; col <= rangeInfo.Address._toCol; col++)
                {
                    var candidate = rangeInfo.GetValue(row, col);
                    if (Evaluate(candidate, searched))
                    {
                        result.Add(internalIndex);
                    }
                    internalIndex++;
                }
            }
            return result;
        }
    }
}
