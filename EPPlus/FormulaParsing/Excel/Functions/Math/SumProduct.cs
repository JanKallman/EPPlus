using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class SumProduct : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            double result = 0d;
            List<List<double>> results = new List<List<double>>();
            foreach (var arg in arguments)
            {
                results.Add(new List<double>());
                var currentResult = results.Last();
                if (arg.Value is IEnumerable<FunctionArgument>)
                {
                    foreach (var val in (IEnumerable<FunctionArgument>)arg.Value)
                    {
                        AddValue(val.Value, currentResult);
                    }
                }
                else if (arg.IsExcelRange)
                {
                    foreach (var val in arg.ValueAsRangeInfo)
                    {
                        AddValue(val.Value, currentResult);
                    }
                }
            }
            // Validate that all supplied lists have the same length
            var arrayLength = results.First().Count;
            foreach (var list in results)
            {
                if (list.Count != arrayLength)
                {
                    throw new ExcelFunctionException("All supplied arrays must have the same length", ExcelErrorCodes.Value);
                }
            }
            for (var rowIndex = 0; rowIndex < arrayLength; rowIndex++)
            {
                double rowResult = 1;
                for (var colIndex = 0; colIndex < results.Count; colIndex++)
                {
                    rowResult *= results[colIndex][rowIndex];
                }
                result += rowResult;
            }
            return CreateResult(result, DataType.Decimal);
        }

        private void AddValue(object convertVal, List<double> currentResult)
        {
            if (IsNumeric(convertVal))
            {
                currentResult.Add(Convert.ToDouble(convertVal));
            }
            else
            {
                currentResult.Add(0d);
            }
        }
    }
}
