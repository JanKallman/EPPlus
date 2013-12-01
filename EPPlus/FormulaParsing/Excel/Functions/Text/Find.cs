using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Find : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
            ValidateArguments(functionArguments, 2);
            var search = ArgToString(functionArguments, 0);
            var searchIn = ArgToString(functionArguments, 1);
            var startIndex = 0;
            if (functionArguments.Count() > 2)
            {
                startIndex = ArgToInt(functionArguments, 2);
            }
            var result = searchIn.IndexOf(search, startIndex, System.StringComparison.Ordinal);
            if (result == -1)
            {
                throw new ExcelFunctionException("Searched phrase was not found", ExcelErrorCodes.Value);
            }
            // Adding 1 because Excel uses 1-based index
            return CreateResult(result + 1, DataType.Integer);
        }
    }
}
