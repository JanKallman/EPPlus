using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    public class ExcelCalculationOption
    {
        public ExcelCalculationOption()
        {
            AllowCirculareReferences = false;
        }
        public bool AllowCirculareReferences { get; set; }
    }
}
