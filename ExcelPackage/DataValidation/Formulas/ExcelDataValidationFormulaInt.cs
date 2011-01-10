using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Globalization;

namespace OfficeOpenXml.DataValidation.Formulas
{
    internal class ExcelDataValidationFormulaInt : ExcelDataValidationFormulaValue<int?>, IExcelDataValidationFormulaInt
    {
        public ExcelDataValidationFormulaInt(XmlNamespaceManager namespaceManager, XmlNode topNode, string formulaPath)
            : base(namespaceManager, topNode, formulaPath)
        {
            var value = GetXmlNodeString(formulaPath);
            if (!string.IsNullOrEmpty(value))
            {
                int intValue = default(int);
                if (int.TryParse(value, out intValue))
                {
                    Value = intValue;
                }
                else
                {
                    ExcelFormula = value;
                }
            }
        }

        protected override string GetValueAsString()
        {
            return Value.HasValue ? Value.Value.ToString() : string.Empty;
        }
    }
}
