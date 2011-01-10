using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Formulas;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Data validation for integer values.
    /// </summary>
    public class ExcelDataValidationInt : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationInt(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : base(worksheet, address, validationType)
        {
            Formula = new ExcelDataValidationFormulaInt(worksheet.NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaInt(worksheet.NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        internal ExcelDataValidationInt(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet, address, validationType, itemElementNode)
        {
            Formula = new ExcelDataValidationFormulaInt(worksheet.NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaInt(worksheet.NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        /// <param name="namespaceManager">For test purposes</param>
        public ExcelDataValidationInt(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
            : base(worksheet, address, validationType, itemElementNode, namespaceManager)
        {
            Formula = new ExcelDataValidationFormulaInt(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaInt(NameSpaceManager, TopNode, _formula2Path);
        }
    }
}
