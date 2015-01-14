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
 * Mats Alm   		                Added		                2015-01-11
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.XPath;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class CountIfs : ExcelFunction
    {
        private readonly NumericExpressionEvaluator _numericExpressionEvaluator;
        private readonly WildCardValueMatcher _wildCardValueMatcher;

         public CountIfs()
            : this(new NumericExpressionEvaluator(), new WildCardValueMatcher())
        {

        }

        public CountIfs(NumericExpressionEvaluator evaluator, WildCardValueMatcher wildCardValueMatcher)
        {
            Utilities.Require.That(evaluator).Named("evaluator").IsNotNull();
            Require.That(wildCardValueMatcher).Named("wildCardValueMatcher").IsNotNull();
            _numericExpressionEvaluator = evaluator;
            _wildCardValueMatcher = wildCardValueMatcher;
        }

        private bool Evaluate(object obj, object expression)
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

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 2);
            var argRanges = new List<ExcelDataProvider.IRangeInfo>();
            var criterias = new List<object>();
            for (var ix = 0; ix < 30; ix +=2)
            {
                if (functionArguments.Length <= ix) break;
                var rangeInfo = functionArguments[ix].ValueAsRangeInfo;
                argRanges.Add(rangeInfo);
                if (ix > 0)
                {
                    ThrowExcelErrorValueExceptionIf(() => rangeInfo.GetNCells() != argRanges[0].GetNCells(), eErrorType.Value);
                }
                criterias.Add(functionArguments[ix+1].Value);
            }
            IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0]);
            var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
            for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
            {
                var indexes = GetMatchIndexes(argRanges[ix], criterias[ix]);
                matchIndexes = enumerable.Intersect(indexes);
            }
            
            return CreateResult((double)matchIndexes.Count(), DataType.Integer);
        }

        private List<int> GetMatchIndexes(ExcelDataProvider.IRangeInfo rangeInfo, object searched)
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
                        result.Add(internalIndex++);
                    }
                }
            }
            return result;
        }
    }
}
