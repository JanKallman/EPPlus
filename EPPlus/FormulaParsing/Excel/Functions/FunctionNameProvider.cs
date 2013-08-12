using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class FunctionNameProvider : IFunctionNameProvider
    {
        private FunctionNameProvider()
        {

        }

        public static FunctionNameProvider Empty
        {
            get { return new FunctionNameProvider(); }
        }

        public virtual bool IsFunctionName(string name)
        {
            return false;
        }
    }
}
