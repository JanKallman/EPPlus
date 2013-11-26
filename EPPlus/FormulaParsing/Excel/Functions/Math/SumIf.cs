using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class SumIf : HiddenValuesHandlingFunction
    {
        private readonly NumericExpressionEvaluator _evaluator;

        public SumIf()
            : this(new NumericExpressionEvaluator())
        {

        }

        public SumIf(NumericExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            _evaluator = evaluator;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var args = arguments.ElementAt(0).Value as IEnumerable<FunctionArgument>;
            var criteria = arguments.ElementAt(1).Value;
            ThrowExcelFunctionExceptionIf(() => criteria == null || criteria.ToString().Length > 255, ExcelErrorCodes.Value);
            var retVal = 0d;
            if (arguments.Count() > 2)
            {
                var sumRange = arguments.ElementAt(2).Value as IEnumerable<FunctionArgument>;
                retVal = CalculateWithSumRange(args, criteria.ToString(), sumRange, context);
            }
            else
            {
                retVal = CalculateSingleRange(args, criteria.ToString());
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double CalculateWithSumRange(IEnumerable<FunctionArgument> range, string criteria, IEnumerable<FunctionArgument> sumRange, ParsingContext context)
        {
            var retVal = 0d;
            var flattenedRange = ArgsToDoubleEnumerable(range, context);
            var flattenedSumRange = ArgsToDoubleEnumerable(sumRange, context);
            for (var x = 0; x < flattenedRange.Count(); x++)
            {
                var candidate = flattenedSumRange.ElementAt(x);
                if (_evaluator.Evaluate(flattenedRange.ElementAt(x), criteria))
                {
                    retVal += Convert.ToDouble(candidate);
                }
            }
            return retVal;
        }

        private double CalculateSingleRange(IEnumerable<FunctionArgument> args, string expression)
        {
            var retVal = 0d;
            if (args != null)
            {
                foreach (var arg in args)
                {
                    retVal += Calculate(arg, expression);
                }
            }
            return retVal;
        }

        private double Calculate(FunctionArgument arg, string expression)
        {
            var retVal = 0d;
            if (ShouldIgnore(arg) || !_evaluator.Evaluate(arg.Value, expression))
            {
                return retVal;
            }
            if (arg.Value is double || arg.Value is int)
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
