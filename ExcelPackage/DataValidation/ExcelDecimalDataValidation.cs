using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Data validation for decimal values
    /// </summary>
    public class ExcelDecimalDataValidation : ExcelDataValidation
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        public ExcelDecimalDataValidation(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : base(worksheet, address, validationType)
        {

        }

        private decimal? GetAsDecimal(string path)
        {
            var val = GetXmlNodeString(path);
            if (string.IsNullOrEmpty(val))
            {
                return null;
            }
            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            if (decimalSeparator == ",")
            {
                val = val.Replace('.', ',');
            }
            return decimal.Parse(val);
        }

        /// <summary>
        /// Corresponds to Formula1 in validation
        /// </summary>
        public decimal? Value
        {
            get
            {
                return GetAsDecimal(_formula1Path);
            }
            set
            {
                SetValue<decimal>(value, _formula1Path);
            }
        }

        /// <summary>
        /// Corresponds to Formula2 in validation
        /// </summary>
        public decimal? Value2
        {
            get
            {
                return GetAsDecimal(_formula2Path);
            }
            set
            {
                SetValue<decimal>(value, _formula2Path);
            }
        }
    }
}
