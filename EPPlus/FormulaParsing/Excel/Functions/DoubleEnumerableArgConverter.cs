using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class DoubleEnumerableArgConverter : CollectionFlattener<double>
    {
        public virtual IEnumerable<double> ConvertArgs(IEnumerable<FunctionArgument> arguments)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
                {
                    var cellInfo = arg.Value as EpplusExcelDataProvider.CellInfo;
                    var value = cellInfo != null ? cellInfo.Value : arg.Value;
                    if (value is double || value is int)
                    {
                        argList.Add(Convert.ToDouble(value));
                    }
            });
        }

        public virtual IEnumerable<double> ConvertArgsIncludingOtherTypes(IEnumerable<FunctionArgument> arguments)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
            {
                var cellInfo = arg.Value as EpplusExcelDataProvider.CellInfo;
                var value = cellInfo != null ? cellInfo.Value : arg.Value;
                if (value is double || value is int || value is bool)
                {
                    argList.Add(Convert.ToDouble(value));
                }
                else if (value is string)
                {
                    argList.Add(0d);
                }
            });
        }
    }
}
