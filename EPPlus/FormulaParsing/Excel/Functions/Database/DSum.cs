using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    public class Dsum : DatabaseFunction
    {
        public Dsum()
            : this(new RowMatcher())
        {

        }

        public Dsum(RowMatcher rowMatcher)
            : base(rowMatcher)
        {

        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var values = GetMatchingValues(arguments, context);
            if (!values.Any()) return CreateResult(0d, DataType.Integer);
            return CreateResult(values.Sum(), DataType.Integer);
        }
    }
}
