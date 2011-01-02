using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation
{
    public class ExcelIntDataValidation : ExcelDataValidation
    {
        public ExcelIntDataValidation(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : base(worksheet, address, validationType)
        {

        }

        private int? GetAsInt(string path)
        {
            var val = GetXmlNodeString(path);
            if (string.IsNullOrEmpty(val))
            {
                return null;
            }
            return int.Parse(val);
        }

        /// <summary>
        /// Corresponds to Formula1 in validation
        /// </summary>
        public int? Value
        {
            get
            {
                return GetAsInt(_formula1Path);
            }
            set
            {
                SetValue<int>(value, _formula1Path);
            }
        }

        public int? Value2
        {
            get
            {
                return GetAsInt(_formula2Path);
            }
            set
            {
                SetValue<int>(value, _formula2Path);
            }
        }
    }
}
