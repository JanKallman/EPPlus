using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    public class ExcelProtectedRange
    {
        public string Name { get; set; }
        public ExcelAddress Address { get; set; }

        public ExcelProtectedRange(string name, ExcelAddress address)
        {
            Name = name;
            Address = address;
        }
    }
}
