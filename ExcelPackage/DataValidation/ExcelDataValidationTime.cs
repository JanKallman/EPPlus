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
    /// Validation for times (<see cref="ExcelTime"/>).
    /// </summary>
    public class ExcelDataValidationTime : ExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime>
    {
         /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : base(worksheet, address, validationType)
        {
            Formula = new ExcelDataValidationFormulaTime(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaTime(NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        internal ExcelDataValidationTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet, address, validationType, itemElementNode)
        {
            Formula = new ExcelDataValidationFormulaTime(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaTime(NameSpaceManager, TopNode, _formula2Path);
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        /// <param name="itemElementNode"></param>
        public ExcelDataValidationTime(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
            : base(worksheet, address, validationType, itemElementNode, namespaceManager)
        {
            Formula = new ExcelDataValidationFormulaTime(NameSpaceManager, TopNode, _formula1Path);
            Formula2 = new ExcelDataValidationFormulaTime(NameSpaceManager, TopNode, _formula2Path);
        }
    }
}
