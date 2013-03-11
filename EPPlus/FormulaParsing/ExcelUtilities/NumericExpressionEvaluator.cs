/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class NumericExpressionEvaluator
    {
        private ValueMatcher _valueMatcher;
        private CompileResultFactory _compileResultFactory;

        public NumericExpressionEvaluator()
            : this(new ValueMatcher(), new CompileResultFactory())
        {

        }

        public NumericExpressionEvaluator(ValueMatcher valueMatcher, CompileResultFactory compileResultFactory)
        {
            _valueMatcher = valueMatcher;
            _compileResultFactory = compileResultFactory;
        }

        private string GetNonNumericStartChars(string expression)
        {
            if (!string.IsNullOrEmpty(expression))
            {
                if (Regex.IsMatch(expression, @"^([^\d]{2})")) return expression.Substring(0, 2);
                if (Regex.IsMatch(expression, @"^([^\d]{1})")) return expression.Substring(0, 1);
            }
            return null;
        }

        public double? OperandAsDouble(object op)
        {
            if (op is double || op is int)
            {
                return Convert.ToDouble(op);
            }
            if (op != null)
            {
                double output;
                if (double.TryParse(op.ToString(), out output))
                {
                    return output;
                }
            }
            return null;
        }

        public bool Evaluate(object left, string expression)
        {
            var operatorCandidate = GetNonNumericStartChars(expression);
            var leftNum = OperandAsDouble(left);
            if (!string.IsNullOrEmpty(operatorCandidate) && leftNum != null)
            {
                IOperator op;
                if (OperatorsDict.Instance.TryGetValue(operatorCandidate, out op))
                {
                    var numericCandidate = expression.Replace(operatorCandidate, string.Empty);
                    double d;
                    if (double.TryParse(numericCandidate, out d))
                    {
                        var leftResult = _compileResultFactory.Create(leftNum);
                        var rightResult = _compileResultFactory.Create(d);
                        var result = op.Apply(leftResult, rightResult);
                        if (result.DataType != DataType.Boolean)
                        {
                            throw new ArgumentException("Illegal operator in expression");
                        }
                        return (bool)result.Result;
                    }
                }
            }
            return _valueMatcher.IsMatch(left, expression) == 0;
        }
    }
}
