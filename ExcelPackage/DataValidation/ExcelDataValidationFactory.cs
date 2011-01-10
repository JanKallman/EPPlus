using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Factory class for ExcelDataValidation.
    /// </summary>
    internal static class ExcelDataValidationFactory
    {
        /// <summary>
        /// Creates an instance of <see cref="ExcelDataValidation"/> out of the given parameters.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="itemElementNode"></param>
        /// <returns></returns>
        public static ExcelDataValidation Create(ExcelDataValidationType type, ExcelWorksheet worksheet, string address, XmlNode itemElementNode)
        {
            Require.Argument(type).IsNotNull("validationType");
            switch (type.ValidationType)
            {
                case eDataValidationType.TextLength:
                case eDataValidationType.Whole:
                    return new ExcelDataValidationInt(worksheet, address, type, itemElementNode);
                case eDataValidationType.Decimal:
                    return new ExcelDataValidationDecimal(worksheet, address, type, itemElementNode);
                default:
                    throw new InvalidOperationException("Non supported validationtype: " + type.ValidationType.ToString());
            }
        }
    }
}
