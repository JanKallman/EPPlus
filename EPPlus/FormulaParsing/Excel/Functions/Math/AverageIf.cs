using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class AverageIf : HiddenValuesHandlingFunction
    {
        private readonly NumericExpressionEvaluator _numericExpressionEvaluator;
        private readonly WildCardValueMatcher _wildCardValueMatcher;

        public AverageIf()
            : this(new NumericExpressionEvaluator(), new WildCardValueMatcher())
        {

        }

        public AverageIf(NumericExpressionEvaluator evaluator, WildCardValueMatcher wildCardValueMatcher)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            Require.That(evaluator).Named("wildCardValueMatcher").IsNotNull();
            _numericExpressionEvaluator = evaluator;
            _wildCardValueMatcher = wildCardValueMatcher;
        }

        private bool Evaluate(object obj, string expression)
        {
            double? candidate = default(double?);
            if (IsNumber(obj))
            {
                candidate = Convert.ToDouble(obj);
            }
            else if (obj is System.DateTime)
            {
                candidate = ((System.DateTime) obj).ToOADate();
            }
            if (candidate.HasValue)
            {
                return _numericExpressionEvaluator.Evaluate(candidate.Value, expression);
            }
            else
            {
                return _wildCardValueMatcher.IsMatch(expression, obj.ToString()) == 0;
            }
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var firstArg = arguments.ElementAt(0);
            var args = firstArg.Value as IEnumerable<FunctionArgument>;
            if (args == null && firstArg.IsExcelRange)
            {
                args = new List<FunctionArgument>(){ firstArg };
            }
            var criteria = arguments.ElementAt(1).Value;
            ThrowExcelFunctionExceptionIf(() => criteria == null || criteria.ToString().Length > 255, ExcelErrorCodes.Value);
            var retVal = 0d;
            if (arguments.Count() > 2)
            {
                var secondArg = arguments.ElementAt(2);
                var lookupRange = secondArg.Value as IEnumerable<FunctionArgument>;
                if (lookupRange == null && secondArg.IsExcelRange)
                {
                    lookupRange = new List<FunctionArgument>() {secondArg};
                }
                retVal = CalculateWithLookupRange(args, criteria.ToString(), lookupRange, context);
            }
            else
            {
                retVal = CalculateSingleRange(args, criteria.ToString(), context);
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double CalculateWithLookupRange(IEnumerable<FunctionArgument> range, string criteria, IEnumerable<FunctionArgument> sumRange, ParsingContext context)
        {
            var retVal = 0d;
            var nMatches = 0;
            var flattenedRange = ArgsToObjectEnumerable(false, range, context);
            var flattenedSumRange = ArgsToDoubleEnumerable(sumRange, context);
            for (var x = 0; x < flattenedRange.Count(); x++)
            {
                var candidate = flattenedSumRange.ElementAt(x);
                if (Evaluate(flattenedRange.ElementAt(x), criteria))
                {
                    nMatches++;
                    retVal += candidate;
                }
            }
            return retVal / nMatches;
        }

        private double CalculateSingleRange(IEnumerable<FunctionArgument> args, string expression, ParsingContext context)
        {
            var retVal = 0d;
            var nMatches = 0;
            var flattendedRange = ArgsToDoubleEnumerable(args, context);
            var candidates = flattendedRange as double[] ?? flattendedRange.ToArray();
            foreach (var candidate in candidates)
            {
                if (Evaluate(candidate, expression))
                {
                    retVal += candidate;
                    nMatches++;
                }
            }
            return retVal / nMatches;
        }

        private double Calculate(FunctionArgument arg, string expression)
        {
            var retVal = 0d;
            if (ShouldIgnore(arg) || !_numericExpressionEvaluator.Evaluate(arg.Value, expression))
            {
                return retVal;
            }
            if (IsNumber(arg.Value))
            {
                retVal += Convert.ToDouble(arg.Value);
            }
            else if (arg.Value is System.DateTime)
            {
                retVal += Convert.ToDateTime(arg.Value).ToOADate();
            }
            else if (arg.Value is IEnumerable<FunctionArgument>)
            {
                foreach (var item in (IEnumerable<FunctionArgument>)arg.Value)
                {
                    retVal += Calculate(item, expression);
                }
            }
            return retVal;
        }
    }
}
