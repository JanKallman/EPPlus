/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class ExpressionEvaluator
    {
        private readonly WildCardValueMatcher _wildCardValueMatcher;
        private readonly CompileResultFactory _compileResultFactory;

        public ExpressionEvaluator()
            : this(new WildCardValueMatcher(), new CompileResultFactory())
        {

        }

        public ExpressionEvaluator(WildCardValueMatcher wildCardValueMatcher, CompileResultFactory compileResultFactory)
        {
            _wildCardValueMatcher = wildCardValueMatcher;
            _compileResultFactory = compileResultFactory;
        }

        private string GetNonAlphanumericStartChars(string expression)
        {
            if (!string.IsNullOrEmpty(expression))
            {
                if (Regex.IsMatch(expression, @"^([^a-zA-Z0-9]{2})")) return expression.Substring(0, 2);
                if (Regex.IsMatch(expression, @"^([^a-zA-Z0-9]{1})")) return expression.Substring(0, 1);
            }
            return null;
        }

        private bool EvaluateOperator(object left, object right, IOperator op)
        {
            var leftResult = _compileResultFactory.Create(left);
            var rightResult = _compileResultFactory.Create(right);
            var result = op.Apply(leftResult, rightResult);
            if (result.DataType != DataType.Boolean)
            {
                throw new ArgumentException("Illegal operator in expression");
            }
            return (bool)result.Result;
        }

        public bool TryConvertToDouble(object op, out double d)
        {
            if (op is double || op is int)
            {
                d = Convert.ToDouble(op);
                return true;
            }
            else if (op is DateTime)
            {
                d = ((DateTime) op).ToOADate();
                return true;
            }
            else if (op != null)
            {
                if (double.TryParse(op.ToString(), out d))
                {
                    return true;
                }
            }
            d = 0;
            return false;
        }

        public bool Evaluate(object left, string expression)
        {
            if (expression == string.Empty)
            {
                return left == null;
            }
            var operatorCandidate = GetNonAlphanumericStartChars(expression);
            if (!string.IsNullOrEmpty(operatorCandidate) && operatorCandidate != "-")
            {
                IOperator op;
                if (OperatorsDict.Instance.TryGetValue(operatorCandidate, out op))
                {
                    var right = expression.Replace(operatorCandidate, string.Empty);
                    if (left == null && right == string.Empty)
                    {
                        return op.Operator == Operators.Equals;
                    }
                    if (left == null ^ right == string.Empty)
                    {
                        return op.Operator == Operators.NotEqualTo;
                    }
                    double leftNum, rightNum;
                    DateTime date;
                    bool leftIsNumeric = TryConvertToDouble(left, out leftNum);
                    bool rightIsNumeric = double.TryParse(right, out rightNum);
                    bool rightIsDate = DateTime.TryParse(right, out date);
                    if (leftIsNumeric && rightIsNumeric)
                    {
                         return EvaluateOperator(leftNum, rightNum, op);
                    }
                    if (leftIsNumeric && rightIsDate)
                    {
                        return EvaluateOperator(leftNum, date.ToOADate(), op);
                    }
                    if (leftIsNumeric != rightIsNumeric)
                    {
                        return op.Operator == Operators.NotEqualTo;
                    }
                    return EvaluateOperator(left, right, op);
                }
            }
            return _wildCardValueMatcher.IsMatch(expression, left) == 0;
        }
    }
}
