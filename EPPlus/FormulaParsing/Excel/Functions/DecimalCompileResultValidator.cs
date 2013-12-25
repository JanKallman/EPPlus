using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class DecimalCompileResultValidator : CompileResultValidator
    {
        public override void Validate(object obj)
        {
            var num = ConvertUtil.GetValueDouble(obj);
            if (double.IsNaN(num) || double.IsInfinity(num))
            {
                throw new ExcelErrorValueException(eErrorType.Num);
            }
        }
    }
}
