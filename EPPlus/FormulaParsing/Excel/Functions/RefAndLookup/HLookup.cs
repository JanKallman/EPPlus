using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class HLookup : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 3);
            var lookupArgs = new LookupArguments(arguments);
            ThrowExcelFunctionExceptionIf(() => lookupArgs.LookupIndex < 1, ExcelErrorCodes.Value);
            var navigator = new LookupNavigator(LookupDirection.Horizontal, lookupArgs, context);
            return Lookup(navigator, lookupArgs);
        }
    }
}
