using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public abstract class CollectionFlattener<T>
    {
        public virtual IEnumerable<T> FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments, Action<FunctionArgument, IList<T>> convertFunc)
        {
            var argList = new List<T>();
            FuncArgsToFlatEnumerable(arguments, argList, convertFunc);
            return argList;
        }

        private void FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments, List<T> argList, Action<FunctionArgument, IList<T>> convertFunc)
        {
            foreach (var arg in arguments)
            {
                if (arg.Value is IEnumerable<FunctionArgument>)
                {
                    FuncArgsToFlatEnumerable((IEnumerable<FunctionArgument>)arg.Value, argList, convertFunc);
                }
                else
                {
                    convertFunc(arg, argList);
                }
            }
        }
    }
}
