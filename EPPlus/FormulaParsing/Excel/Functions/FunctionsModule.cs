using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public abstract class FunctionsModule : IFunctionModule
    {
        private readonly Dictionary<string, ExcelFunction> _functions = new Dictionary<string, ExcelFunction>();

        public IDictionary<string, ExcelFunction> Functions
        {
            get { return _functions; }
        }
    }
}
