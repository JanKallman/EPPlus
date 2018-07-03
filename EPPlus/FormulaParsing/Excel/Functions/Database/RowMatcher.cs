﻿/* Copyright (C) 2011  Jan Källman
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
 * Mats Alm   		                Added		                2015-04-06
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    public class RowMatcher
    {
        private readonly WildCardValueMatcher _wildCardValueMatcher;
        private readonly ExpressionEvaluator _expressionEvaluator;

        public RowMatcher()
            : this(new WildCardValueMatcher(), new ExpressionEvaluator())
        {
            
        }

        public RowMatcher(WildCardValueMatcher wildCardValueMatcher, ExpressionEvaluator expressionEvaluator)
        {
            _wildCardValueMatcher = wildCardValueMatcher;
            _expressionEvaluator = expressionEvaluator;
        }

        public bool IsMatch(ExcelDatabaseRow row, ExcelDatabaseCriteria criteria)
        {
            var retVal = true;
            foreach (var c in criteria.Items)
            {
                var candidate = c.Key.FieldIndex.HasValue ? row[c.Key.FieldIndex.Value] : row[c.Key.FieldName];
                var crit = c.Value;
                if (candidate.IsNumeric() && crit.IsNumeric())
                {
                    if(System.Math.Abs(ConvertUtil.GetValueDouble(candidate) - ConvertUtil.GetValueDouble(crit)) > double.Epsilon) return false;
                }
                else
                {
                    var criteriaString = crit.ToString();
                    if (!Evaluate(candidate, criteriaString))
                    {
                        return false;
                    }
                }
            }
            return retVal;
        }

        private bool Evaluate(object obj, string expression)
        {
            if (obj == null) return false;
            double? candidate = default(double?);
            if (ConvertUtil.IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue)
            {
                return _expressionEvaluator.Evaluate(candidate.Value, expression);
            }
            return _wildCardValueMatcher.IsMatch(expression, obj.ToString()) == 0;
        }
    }
}
