using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    public class Dmin : DatabaseFunction
    {
        public Dmin()
            : this(new RowMatcher())
        {

        }

        public Dmin(RowMatcher rowMatcher)
            : base(rowMatcher)
        {

        }
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var values = GetMatchingValues(arguments, context);
            if (!values.Any()) return CreateResult(0d, DataType.Integer);
            return CreateResult(values.Min(), DataType.Integer);
        }
    }
}
