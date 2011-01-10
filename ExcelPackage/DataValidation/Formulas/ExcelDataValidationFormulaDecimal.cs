using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Globalization;

namespace OfficeOpenXml.DataValidation.Formulas
{
    /// <summary>
    /// 
    /// </summary>
    internal class ExcelDataValidationFormulaDecimal : ExcelDataValidationFormulaValue<double?>, IExcelDataValidationFormulaDecimal
    {
        public ExcelDataValidationFormulaDecimal(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
            : base(namespaceManager, topNode, formulaPath)
        {
            var value = GetXmlNodeString(formulaPath);
            if (!string.IsNullOrEmpty(value))
            {
                double dValue = default(double);
                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out dValue))
                {
                    Value = dValue;
                }
                else
                {
                    ExcelFormula = value;
                }
            }
        }

        protected override string GetValueAsString()
        {
            return Value.HasValue ? Value.Value.ToString("g15", CultureInfo.InvariantCulture) : string.Empty;
        }
    }
}
