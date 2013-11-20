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
                    if (arg.Value is ExcelDataProvider.IRangeInfo)
                    {
                        foreach (var cell in (ExcelDataProvider.IRangeInfo)arg.Value)
                        {
                            argList.Add(cell.ValueDouble);
                        }
                    }
                    else
                    {
                        if (arg.Value is double || arg.Value is int)
                        {
                            argList.Add(Convert.ToDouble(arg.Value));
                        }
                    }
                });
        }

        public virtual IEnumerable<double> ConvertArgsIncludingOtherTypes(IEnumerable<FunctionArgument> arguments)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
            {
                //var cellInfo = arg.Value as EpplusExcelDataProvider.CellInfo;
                //var value = cellInfo != null ? cellInfo.Value : arg.Value;
                if (arg.Value is ExcelDataProvider.IRangeInfo)
                {
                    foreach (var cell in (ExcelDataProvider.IRangeInfo)arg.Value)
                    {
                        argList.Add(cell.ValueDoubleLogical);
                    }
                }
                else
                {
                    if (arg.Value is double || arg.Value is int || arg.Value is bool)
                    {
                        argList.Add(Convert.ToDouble(arg.Value));
                    }
                    else if (arg.Value is string)
                    {
                        argList.Add(0d);
                    }
                }
            });
        }
    }
}
