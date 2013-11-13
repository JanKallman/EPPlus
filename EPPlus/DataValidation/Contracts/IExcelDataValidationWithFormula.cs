using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Formulas.Contracts;

namespace OfficeOpenXml.DataValidation.Contracts
{
    public interface IExcelDataValidationWithFormula<T> : IExcelDataValidation
        where T : IExcelDataValidationFormula
    {
        T Formula { get; }
    }
}
