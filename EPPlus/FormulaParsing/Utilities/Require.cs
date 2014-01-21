using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Utilities
{
    public static class Require
    {
        public static ArgumentInfo<T> That<T>(T arg)
        {
            return new ArgumentInfo<T>(arg);
        }
    }
}
