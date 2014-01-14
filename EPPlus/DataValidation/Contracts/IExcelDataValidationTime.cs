using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Contracts
{
    /// <summary>
    /// Data validation interface for time validation.
    /// </summary>
    public interface IExcelDataValidationTime : IExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime>, IExcelDataValidationWithOperator
    {
    }
}
