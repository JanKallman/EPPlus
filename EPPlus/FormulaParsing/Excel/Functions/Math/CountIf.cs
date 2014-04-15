using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class CountIf : ExcelFunction
    {
        private readonly NumericExpressionEvaluator _numericExpressionEvaluator;
        private readonly WildCardValueMatcher _wildCardValueMatcher;

        public CountIf()
            : this(new NumericExpressionEvaluator(), new WildCardValueMatcher())
        {

        }

        public CountIf(NumericExpressionEvaluator evaluator, WildCardValueMatcher wildCardValueMatcher)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            Require.That(wildCardValueMatcher).Named("wildCardValueMatcher").IsNotNull();
            _numericExpressionEvaluator = evaluator;
            _wildCardValueMatcher = wildCardValueMatcher;
        }

        private bool Evaluate(object obj, string expression)
        {
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue)
            {
                return _numericExpressionEvaluator.Evaluate(candidate.Value, expression);
            }
            return _wildCardValueMatcher.IsMatch(expression, obj.ToString()) == 0;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 2);
            var range = functionArguments.ElementAt(0);
            var criteria = ArgToString(functionArguments, 1);
            double result = 0d;
            if (range.IsExcelRange)
            {
                foreach (var cell in range.ValueAsRangeInfo)
                {
                    if (Evaluate(cell.Value, criteria))
                    {
                        result++;
                    }
                }
            }
            else if (range.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var arg in (IEnumerable<FunctionArgument>) range.Value)
                {
                    if(Evaluate(arg.Value, criteria))
                    {
                        result++;
                    }
                }
            }
            else
            {
                if (Evaluate(range.Value, criteria))
                {
                    result++;
                }
            }
            return CreateResult(result, DataType.Integer);
        }
    }
}
