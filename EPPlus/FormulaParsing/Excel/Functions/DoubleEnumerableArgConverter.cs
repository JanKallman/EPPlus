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
                if (arg.Value is double || arg.Value is int)
                {
                    argList.Add(Convert.ToDouble(arg.Value));
                }
            });
        }

        public virtual IEnumerable<double> ConvertArgsIncludingOtherTypes(IEnumerable<FunctionArgument> arguments)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
            {
                if (arg.Value is double || arg.Value is int || arg.Value is bool)
                {
                    argList.Add(Convert.ToDouble(arg.Value));
                }
                else if (arg.Value is string)
                {
                    argList.Add(0d);
                }
            });
        }
    }
}
