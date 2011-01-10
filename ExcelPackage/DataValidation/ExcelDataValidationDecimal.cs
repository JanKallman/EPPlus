using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.DataValidation.Formulas;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Data validation for decimal values
    /// </summary>
    public class ExcelDataValidationDecimal : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationDecimal(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : base(worksheet, address, validationType)
        {
            Formula = new ExcelDataValidationFormulaDecimal(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaDecimal(NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        internal ExcelDataValidationDecimal(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet, address, validationType, itemElementNode)
        {
            Formula = new ExcelDataValidationFormulaDecimal(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaDecimal(NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        /// <param name="namespaceManager">For test purposes</param>
        public ExcelDataValidationDecimal(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
            : base(worksheet, address, validationType, itemElementNode, namespaceManager)
        {
            Formula = new ExcelDataValidationFormulaDecimal(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaDecimal(NameSpaceManager, TopNode, _formula2Path);
        }
    }
}
