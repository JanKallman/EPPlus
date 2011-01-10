using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation.Formulas.Contracts
{
    /// <summary>
    /// Interface for a data validation formula of <see cref="System.Decimal"/> value
    /// </summary>
    public interface IExcelDataValidationFormulaDecimal : IExcelDataValidationFormulaWithValue<double?>
    {
    }
}
