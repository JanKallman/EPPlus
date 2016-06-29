using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public static class LookupNavigatorFactory
    {
        public static LookupNavigator Create(LookupDirection direction, LookupArguments args, ParsingContext parsingContext)
        {
            if (args.ArgumentDataType == LookupArguments.LookupArgumentDataType.ExcelRange)
            {
                return new ExcelLookupNavigator(direction, args, parsingContext);
            }
            else if (args.ArgumentDataType == LookupArguments.LookupArgumentDataType.DataArray)
            {
                return new ArrayLookupNavigator(direction, args, parsingContext);
            }
            throw new NotSupportedException("Invalid argument datatype");
        }
    }
}
