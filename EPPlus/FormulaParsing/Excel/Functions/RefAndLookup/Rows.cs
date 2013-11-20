using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class Rows : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var range = ArgToString(arguments, 0);
            if (ExcelAddressUtil.IsValidAddress(range))
            {
                var factory = new RangeAddressFactory(context.ExcelDataProvider);
                var address = factory.Create(range);
                return CreateResult(address.ToRow - address.FromRow + 1, DataType.Integer);
            }
            throw new ArgumentException("Invalid range supplied");
        }
    }
}
