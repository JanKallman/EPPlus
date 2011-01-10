using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// warning style, controls how Excel will handle invalid changes.
    /// </summary>
    public enum ExcelDataValidationWarningStyle
    {
        /// <summary>
        /// warning style will be excluded
        /// </summary>
        undefined,
        /// <summary>
        /// stop warning style, invalid changes will not be accepted
        /// </summary>
        stop,
        /// <summary>
        /// warning will be presented when an attempt to an invalid change is done, but the change will be accepted.
        /// </summary>
        warning,
        /// <summary>
        /// information warning style.
        /// </summary>
        information
    }
}
