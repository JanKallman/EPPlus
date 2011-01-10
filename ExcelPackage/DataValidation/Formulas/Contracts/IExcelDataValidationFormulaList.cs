using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation.Formulas.Contracts
{
    /// <summary>
    /// Interface for a data validation of list type
    /// </summary>
    public interface IExcelDataValidationFormulaList : IExcelDataValidationFormula
    {
        /// <summary>
        /// A list of value strings.
        /// </summary>
        IList<string> Values { get; }
    }
}
