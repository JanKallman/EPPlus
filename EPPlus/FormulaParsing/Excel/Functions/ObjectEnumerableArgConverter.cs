using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class ObjectEnumerableArgConverter : CollectionFlattener<object>
    {
        public virtual IEnumerable<object> ConvertArgs(bool ignoreHidden, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
                {
                    if (arg.Value is ExcelDataProvider.IRangeInfo)
                    {
                        foreach (var cell in (ExcelDataProvider.IRangeInfo)arg.Value)
                        {
                            if (!CellStateHelper.ShouldIgnore(ignoreHidden, cell, context))
                            {
                                argList.Add(cell.Value);
                            }
                        }
                    }
                    else
                    {
                       argList.Add(arg.Value);
                    }
                })
            ;

        }
    }
}
