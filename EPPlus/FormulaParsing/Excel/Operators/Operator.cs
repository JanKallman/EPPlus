using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

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
        private int _precedence;
        private Operators _operator;

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
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult((Convert.ToDouble(l.Result)) + (Convert.ToDouble(r.Result)), DataType.Integer);
                    }
                    else if (l.IsNumeric && r.IsNumeric)
                    {
                        return new CompileResult((Convert.ToDouble(l.Result)) + (Convert.ToDouble(r.Result)), DataType.Decimal);
                    }
                    return new CompileResult(0, DataType.Integer);
                }); 
            }
        }

        public static IOperator Minus
        {
            get
            {
                return new Operator(Operators.Minus, PrecedenceAddSubtract, (l, r) =>
                {
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult((Convert.ToDouble(l.Result)) - (Convert.ToDouble(r.Result)), DataType.Integer);
                    }
                    else if (l.IsNumeric && r.IsNumeric)
                    {
                        return new CompileResult((Convert.ToDouble(l.Result)) - (Convert.ToDouble(r.Result)), DataType.Decimal);
                    }
                    return new CompileResult(0, DataType.Integer);
                });
            }
        }

        public static IOperator Multiply
        {
            get
            {
                return new Operator(Operators.Multiply, PrecedenceMultiplyDevide, (l, r) =>
                {
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult((Convert.ToDouble(l.Result)) * (Convert.ToDouble(r.Result)), DataType.Integer);
                    }
                    if (l.IsNumeric && r.IsNumeric)
                    {
                        return new CompileResult((Convert.ToDouble(l.Result)) * (Convert.ToDouble(r.Result)), DataType.Decimal);
                    }
                    return new CompileResult(0, DataType.Integer);
                });
            }
        }

        public static IOperator Divide
        {
            get
            {
                return new Operator(Operators.Divide, PrecedenceMultiplyDevide, (l, r) =>
                {
                    var left = Convert.ToDouble(l.Result);
                    var right = Convert.ToDouble(r.Result);
                    if (right == 0d)
                    {
                        throw new DivideByZeroException(string.Format("left: {0}, right: {1}", left, right));
                    }
                    if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
                    {
                        return new CompileResult(left / right, DataType.Integer);
                    }
                    if (l.IsNumeric && r.IsNumeric)
                    {
                        return new CompileResult(left / right, DataType.Decimal);
                    }
                    return new CompileResult(0, DataType.Integer);
                });
            }
        }

        public static IOperator Exp
        {
            get
            {
                return new Operator(Operators.Exponentiation, PrecedenceExp, (l, r) =>
                    {
                        if (l.IsNumeric && r.IsNumeric)
                        {
                            return new CompileResult(Math.Pow(Convert.ToDouble(l.Result), Convert.ToDouble(r.Result)), DataType.Decimal);
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
                        var lStr = l.Result != null ? l.Result.ToString() : string.Empty;
                        var rStr = r.Result != null ? r.Result.ToString() : string.Empty;
                        return new CompileResult(string.Concat(lStr, rStr), DataType.String);
                    });
            }
        }

        public static IOperator Modulus
        {
            get
            {
                return new Operator(Operators.Modulus, PrecedenceModulus, (l, r) =>
                {
                    return new CompileResult(Convert.ToDouble(l.Result) % Convert.ToDouble(r.Result), DataType.Integer); ;
                });
            }
        }

        public static IOperator GreaterThan
        {
            get
            {
                return new Operator(Operators.GreaterThan, PrecedenceComparison, (l, r) =>
                    {
                        if (l.IsNumeric && r.IsNumeric)
                        {
                            return new CompileResult((Convert.ToDouble(l.Result)) > (Convert.ToDouble(r.Result)), DataType.Boolean);
                        }
                        return new CompileResult(false, DataType.Boolean);
                    });
            }
        }

        public static IOperator Eq
        {
            get
            {
                return new Operator(Operators.Equals, PrecedenceComparison, (l, r) =>
                    {
                        return new CompileResult(l.Result.Equals(r.Result), DataType.Boolean);
                    });
            }
        }

        public static IOperator NotEqualsTo
        {
            get
            {
                return new Operator(Operators.NotEqualTo, PrecedenceComparison, (l, r) =>
                    {
                        return new CompileResult(!l.Result.Equals(r.Result), DataType.Boolean);
                    });
            }
        }

        public static IOperator GreaterThanOrEqual
        {
            get
            {
                return new Operator(Operators.GreaterThanOrEqual, PrecedenceComparison, (l, r) =>
                {
                    if (l.IsNumeric && r.IsNumeric)
                    {
                        return new CompileResult((Convert.ToDouble(l.Result)) >= (Convert.ToDouble(r.Result)), DataType.Boolean);
                    }
                    return new CompileResult(false, DataType.Boolean);
                });
            }
        }

        public static IOperator LessThan
        {
            get
            {
                return new Operator(Operators.LessThan, PrecedenceComparison, (l, r) =>
                    {
                        if (l.IsNumeric && r.IsNumeric)
                        {
                            return new CompileResult((Convert.ToDouble(l.Result)) < (Convert.ToDouble(r.Result)), DataType.Boolean);
                        }
                        return new CompileResult(false, DataType.Boolean);
                    });
            }
        }

        public static IOperator LessThanOrEqual
        {
            get
            {
                return new Operator(Operators.LessThanOrEqual, PrecedenceComparison, (l, r) =>
                {
                    if (l.IsNumeric && r.IsNumeric)
                    {
                        return new CompileResult((Convert.ToDouble(l.Result)) <= (Convert.ToDouble(r.Result)), DataType.Boolean);
                    }
                    return new CompileResult(false, DataType.Boolean);
                });
            }
        }
    }
}
