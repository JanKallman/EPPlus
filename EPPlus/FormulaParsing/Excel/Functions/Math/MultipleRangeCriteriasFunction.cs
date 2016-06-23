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

        private readonly ExpressionEvaluator _expressionEvaluator;

        protected MultipleRangeCriteriasFunction()
            :this(new ExpressionEvaluator())
        {
            
        }

        protected MultipleRangeCriteriasFunction(ExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            _expressionEvaluator = evaluator;
        }

        protected bool Evaluate(object obj, string expression)
        {
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue)
            {
                return _expressionEvaluator.Evaluate(candidate.Value, expression);
            }
            return _expressionEvaluator.Evaluate(obj, expression);
        }

        protected List<int> GetMatchIndexes(ExcelDataProvider.IRangeInfo rangeInfo, string searched)
        {
            var result = new List<int>();
            var internalIndex = 0;
            for (var row = rangeInfo.Address._fromRow; row <= rangeInfo.Address._toRow; row++)
            {
                for (var col = rangeInfo.Address._fromCol; col <= rangeInfo.Address._toCol; col++)
                {
                    var candidate = rangeInfo.GetValue(row, col);
                    if (searched != null && Evaluate(candidate, searched))
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
