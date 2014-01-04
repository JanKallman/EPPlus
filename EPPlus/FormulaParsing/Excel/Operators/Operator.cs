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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    public class Operator : IOperator
    {
        private const int PrecedenceExp = 2;
        private const int PrecedenceMultiplyDevide = 3;
        private const int PrecedenceIntegerDivision = 4;
        private const int PrecedenceModulus = 5;
        private const int PrecedenceAddSubtract = 10;
        private const int PrecedenceConcat = 15;
        private const int PrecedenceComparison = 25;

        private Operator() { }

        private Operator(Operators @operator, int precedence, Func<CompileResult, CompileResult, CompileResult> implementation)
        {
            _implementation = implementation;
            _precedence = precedence;
            _operator = @operator;
        }

        private readonly Func<CompileResult, CompileResult, CompileResult> _implementation;
        private readonly int _precedence;
        private readonly Operators _operator;

        int IOperator.Precedence
        {
            get { return _precedence; }
        }

        Operators IOperator.Operator
        {
            get { return _operator; }
        }

        public CompileResult Apply(CompileResult left, CompileResult right)
        {
            return _implementation(left, right);
        }

        public static IOperator Plus
        {
            get
            {
                return new Operator(Operators.Plus, PrecedenceAddSubtract, (l, r) =>
                {
                    l = l ?? new CompileResult(0, DataType.Integer);
                    r = r ?? new CompileResult(0, DataType.Integer);
                    CheckForErrors(l, r);
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(l.ResultNumeric + r.ResultNumeric, DataType.Integer);
                    }
                    else if ((l.IsNumeric || l.IsNumericString) && (r.IsNumeric || r.IsNumericString))
                    {
                        return new CompileResult(l.ResultNumeric + r.ResultNumeric, DataType.Decimal);
                    }
                    throw new ExcelErrorValueException(eErrorType.Value);
                }); 
            }
        }

        public static IOperator Minus
        {
            get
            {
                return new Operator(Operators.Minus, PrecedenceAddSubtract, (l, r) =>
                {
                    l = l ?? new CompileResult(0, DataType.Integer);
                    r = r ?? new CompileResult(0, DataType.Integer);
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(l.ResultNumeric - r.ResultNumeric, DataType.Integer);
                    }
                    else if ((l.IsNumeric || l.IsNumericString) && (r.IsNumeric || r.IsNumericString))
                    {
                        return new CompileResult(l.ResultNumeric - r.ResultNumeric, DataType.Decimal);
                    }
                    throw new ExcelErrorValueException(eErrorType.Value);
                });
            }
        }

        public static IOperator Multiply
        {
            get
            {
                return new Operator(Operators.Multiply, PrecedenceMultiplyDevide, (l, r) =>
                {
                    l = l ?? new CompileResult(0, DataType.Integer);
                    r = r ?? new CompileResult(0, DataType.Integer);
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Integer);
                    }
                    else if ((l.IsNumeric || l.IsNumericString) && (r.IsNumeric || r.IsNumericString))
                    {
                        return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Decimal);
                    }
                    throw new ExcelErrorValueException(eErrorType.Value);
                });
            }
        }

        public static IOperator Divide
        {
            get
            {
                return new Operator(Operators.Divide, PrecedenceMultiplyDevide, (l, r) =>
                {
                    if (!(l.IsNumeric || l.IsNumericString) || !(r.IsNumeric || r.IsNumericString))
                    {
                        throw new ExcelErrorValueException(eErrorType.Value);
                    }
                    var left = l.ResultNumeric;
                    var right = r.ResultNumeric;
                    if (Math.Abs(right - 0d) < double.Epsilon)
                    {
                        throw new ExcelErrorValueException(eErrorType.Div0);
                    }
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(left / right, DataType.Integer);
                    }
                    else if ((l.IsNumeric || l.IsNumericString) && (r.IsNumeric || r.IsNumericString))
                    {
                        return new CompileResult(left / right, DataType.Decimal);
                    }
                    throw new ExcelErrorValueException(eErrorType.Value);
                });
            }
        }

        public static IOperator Exp
        {
            get
            {
                return new Operator(Operators.Exponentiation, PrecedenceExp, (l, r) =>
                    {
                        if (l == null && r == null)
                        {
                            throw new ExcelErrorValueException(eErrorType.Value);
                        }
                        l = l ?? new CompileResult(0, DataType.Integer);
                        r = r ?? new CompileResult(0, DataType.Integer);
                        if (l.IsNumeric && r.IsNumeric)
                        {
                            return new CompileResult(Math.Pow(l.ResultNumeric, r.ResultNumeric), DataType.Decimal);
                        }
                        return new CompileResult(0d, DataType.Decimal);
                    });
            }
        }

        public static IOperator Concat
        {
            get
            {
                return new Operator(Operators.Concat, PrecedenceConcat, (l, r) =>
                    {
                        l = l ?? new CompileResult(string.Empty, DataType.String);
                        r = r ?? new CompileResult(string.Empty, DataType.String);
                        var lStr = l.Result != null ? l.Result.ToString() : string.Empty;
                        var rStr = r.Result != null ? r.Result.ToString() : string.Empty;
                        return new CompileResult(string.Concat(lStr, rStr), DataType.String);
                    });
            }
        }

        public static IOperator GreaterThan
        {
            get
            {
                return new Operator(Operators.GreaterThan, PrecedenceComparison, (l, r) => new CompileResult(Compare(l, r) > 0, DataType.Boolean));
            }
        }

        public static IOperator Eq
        {
            get
            {
                return new Operator(Operators.Equals, PrecedenceComparison, (l, r) => new CompileResult(Compare(l, r) == 0, DataType.Boolean));
            }
        }

        public static IOperator NotEqualsTo
        {
            get
            {
                return new Operator(Operators.Equals, PrecedenceComparison, (l, r) => new CompileResult(Compare(l, r) != 0, DataType.Boolean));
            }
        }

        public static IOperator GreaterThanOrEqual
        {
            get
            {
                return new Operator(Operators.GreaterThan, PrecedenceComparison, (l, r) => new CompileResult(Compare(l, r) >= 0, DataType.Boolean));
            }
        }

        public static IOperator LessThan
        {
            get
            {
                return new Operator(Operators.GreaterThan, PrecedenceComparison, (l, r) => new CompileResult(Compare(l, r) < 0, DataType.Boolean));
            }
        }

        public static IOperator LessThanOrEqual
        {
            get
            {
                return new Operator(Operators.GreaterThan, PrecedenceComparison, (l, r) => new CompileResult(Compare(l, r) <= 0, DataType.Boolean));
            }
        }

        private static int Compare(CompileResult l, CompileResult r)
        {
            CheckForErrors(l, r);
            if (l.Result == null && r.Result == null)
            {
                return 0;
            }
            if (l.Result == null && r.Result != null)
            {
                return -1;
            }
            if (l.Result != null && r.Result == null)
            {
                return 1;
            }
            if (l.IsNumeric && r.IsNumeric)
            {
                var lnum = l.ResultNumeric;
                var rnum = r.ResultNumeric;
                if (Math.Abs(lnum - rnum) < double.Epsilon) return 0;
                return lnum.CompareTo(rnum);
            }
            else
            {
                return CompareString(l.Result, r.Result);
            }
        }

        private static int CompareString(object l, object r)
        {
            var sl = (l ?? "").ToString();
            var sr = (r ?? "").ToString();
            return System.String.Compare(sl, sr, System.StringComparison.Ordinal);
        }

        private static void  CheckForErrors(CompileResult l, CompileResult r)
        {
            if (l.DataType == DataType.ExcelError)
            {
                throw new ExcelErrorValueException((ExcelErrorValue)l.Result);
            }
            if (r.DataType == DataType.ExcelError)
            {
                throw new ExcelErrorValueException((ExcelErrorValue)r.Result);
            }
        }
    }
}
