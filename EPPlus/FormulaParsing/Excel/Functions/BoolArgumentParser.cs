using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class BoolArgumentParser : ArgumentParser
    {
        public override object Parse(object obj)
        {
            if (obj == null) return false;
            if (obj is bool) return (bool)obj;
            if (obj.IsNumeric()) return Convert.ToBoolean(obj);
            bool result;
            if (bool.TryParse(obj.ToString(), out result))
            {
                return result;
            }
            return result;
        }
    }
}
