using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    public class ExcelDataValidationWithFormula2<T> : ExcelDataValidationWithFormula<T>
        where T : IExcelDataValidationFormula
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationWithFormula2(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : this(worksheet, address, validationType, null)
        {

        }

         /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet">Worksheet that owns the validation</param>
        /// <param name="itemElementNode">Xml top node (dataValidations)</param>
        /// <param name="validationType">Data validation type</param>
        /// <param name="address">address for data validation</param>
        internal ExcelDataValidationWithFormula2(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet, address, validationType, itemElementNode)
        {
            
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet">Worksheet that owns the validation</param>
        /// <param name="itemElementNode">Xml top node (dataValidations)</param>
        /// <param name="validationType">Data validation type</param>
        /// <param name="address">address for data validation</param>
        /// <param name="namespaceManager">for test purposes</param>
        internal ExcelDataValidationWithFormula2(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
            : base(worksheet, address, validationType, itemElementNode, namespaceManager)
        {

        }
        /// <summary>
        /// Formula - Either a {T} value or a spreadsheet formula
        /// </summary>
        public T Formula2
        {
            get;
            protected set;
        }
    }
}
