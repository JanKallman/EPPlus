using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation.Formulas.Contracts
{
    /// <summary>
    /// Interface for a formula with a value
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public interface IExcelDataValidationFormulaWithValue<T> : IExcelDataValidationFormula
    {
        /// <summary>
        /// The value.
        /// </summary>
        T Value { get; set; }
    }
}
