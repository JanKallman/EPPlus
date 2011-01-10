using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation.Formulas.Contracts
{
    /// <summary>
    /// Interface for a data validation formula
    /// </summary>
    public interface IExcelDataValidationFormula
    {
        /// <summary>
        /// An excel formula
        /// </summary>
        string ExcelFormula { get; set; }
    }
}
