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
    public class Row : LookupFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments == null || arguments.Count() == 0)
            {
                return CreateResult(context.Scopes.Current.Address.FromRow + 1, DataType.Integer);
            }
            var rangeAddress = ArgToString(arguments, 0);
            if (ExcelAddressUtil.IsValidAddress(rangeAddress))
            {
                var factory = new RangeAddressFactory(context.ExcelDataProvider);
                var address = factory.Create(rangeAddress);
                return CreateResult(address.FromRow + 1, DataType.Integer);
            }
            throw new ArgumentException("An invalid argument was supplied");
        }
    }
}
