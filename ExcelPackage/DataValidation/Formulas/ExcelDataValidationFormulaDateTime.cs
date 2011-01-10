using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Globalization;

namespace OfficeOpenXml.DataValidation.Formulas
{
    internal class ExcelDataValidationFormulaDateTime : ExcelDataValidationFormulaValue<DateTime?>, IExcelDataValidationFormulaDateTime
    {
        public ExcelDataValidationFormulaDateTime(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
            : base(namespaceManager, topNode, formulaPath)
        {
            var value = GetXmlNodeString(formulaPath);
            if (!string.IsNullOrEmpty(value))
            {
                double oADate = default(double);
                if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out oADate))
                {
                    Value = DateTime.FromOADate(oADate);
                }
                else
                {
                    ExcelFormula = value;
                }
            }
        }

        protected override string GetValueAsString()
        {
            return Value.HasValue ? Value.Value.ToOADate().ToString(CultureInfo.InvariantCulture) : string.Empty;
        }

    }
}
