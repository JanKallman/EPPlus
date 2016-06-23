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
        private readonly ExpressionEvaluator _expressionEvaluator;

        public CountIf()
            : this(new ExpressionEvaluator())
        {

        }

        public CountIf(ExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            _expressionEvaluator = evaluator;
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
                return _expressionEvaluator.Evaluate(candidate.Value, expression);
            }
            return _expressionEvaluator.Evaluate(obj, expression);
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 2);
            var range = functionArguments.ElementAt(0);
            var criteria = functionArguments.ElementAt(1).ValueFirst != null ? ArgToString(functionArguments, 1) : null;
            double result = 0d;
            if (range.IsExcelRange)
            {
                ExcelDataProvider.IRangeInfo rangeInfo = range.ValueAsRangeInfo;
                for (int row = rangeInfo.Address.Start.Row; row < rangeInfo.Address.End.Row + 1; row++)
                {
                    for (int col = rangeInfo.Address.Start.Column; col < rangeInfo.Address.End.Column + 1; col++)
                    {
                        if (criteria != null && Evaluate(rangeInfo.Worksheet.GetValue(row, col), criteria))
                        {
                            result++;
                        }
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
