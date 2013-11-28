using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public abstract class ArgumentParser
    {
        public abstract object Parse(object obj);
    }
}
