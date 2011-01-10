using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Operator for comparison between Formula and Formula2 in a validation.
    /// </summary>
    public enum ExcelDataValidationOperator
    {
        any,
        equal,
        notEqual,
        lessThan,
        lessThanOrEqual,
        greaterThan,
        greaterThanOrEqual,
        between,
        notBetween
    }
}
