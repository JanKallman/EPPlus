using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace EPPlusTest.FormulaParsing.TestHelpers
{
    public static class FunctionsHelper
    {
        public static IEnumerable<FunctionArgument> CreateArgs(params object[] args)
        {
            var list = new List<FunctionArgument>();
            foreach (var arg in args)
            {
                list.Add(new FunctionArgument(arg));
            }
            return list;
        }
    }
}
