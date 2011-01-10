using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.DataValidation.Formulas;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Validation for <see cref="DateTime"/>.
    /// </summary>
    public class ExcelDataValidationDateTime : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationDateTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : base(worksheet, address, validationType)
        {
            Formula = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        internal ExcelDataValidationDateTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet, address, validationType, itemElementNode)
        {
            Formula = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        public ExcelDataValidationDateTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
            : base(worksheet, address, validationType, itemElementNode, namespaceManager)
        {
            Formula = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaDateTime(NameSpaceManager, TopNode, _formula2Path);
        }
    }
}
