using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    public abstract class DatabaseFunction : ExcelFunction
    {
        protected RowMatcher RowMatcher { get; private set; }

        public DatabaseFunction()
            : this(new RowMatcher())
        {
            
        }

        public DatabaseFunction(RowMatcher rowMatcher)
        {
            RowMatcher = rowMatcher;
        }

        protected IEnumerable<double> GetMatchingValues(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            var dbAddress = arguments.ElementAt(0).ValueAsRangeInfo.Address.Address;
            var field = ArgToString(arguments, 1).ToLower(CultureInfo.InvariantCulture);
            var criteriaRange = arguments.ElementAt(2).ValueAsRangeInfo.Address.Address;

            var db = new ExcelDatabase(context.ExcelDataProvider, dbAddress);
            var criteria = new ExcelDatabaseCriteria(context.ExcelDataProvider, criteriaRange);
            var values = new List<double>();

            while (db.HasMoreRows)
            {
                var dataRow = db.Read();
                if (!RowMatcher.IsMatch(dataRow, criteria.Items)) continue;
                if (string.IsNullOrEmpty(field)) continue;
                var candidate = dataRow[field];
                if (ConvertUtil.IsNumeric(candidate))
                {
                    values.Add(ConvertUtil.GetValueDouble(candidate));
                }
            }
            return values;
        }
    }
}
