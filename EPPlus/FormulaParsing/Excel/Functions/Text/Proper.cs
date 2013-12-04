using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    public class Proper : ExcelFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var text = ArgToString(arguments, 0).ToLower();
            var sb = new StringBuilder();
            var previousChar = '.';
            foreach (var ch in text)
            {
                if (!char.IsLetter(previousChar))
                {
                    sb.Append(ch.ToString().ToUpperInvariant());
                }
                else
                {
                    sb.Append(ch);
                }
                previousChar = ch;
            }
            return CreateResult(sb.ToString(), DataType.String);
        }
    }
}
