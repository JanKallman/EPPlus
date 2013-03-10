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
