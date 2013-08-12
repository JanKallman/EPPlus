using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation.Contracts
{
    /// <summary>
    /// Interface for data validation
    /// </summary>
    public interface IExcelDataValidation
    {
        /// <summary>
        /// Address of data validation
        /// </summary>
        ExcelAddress Address { get; }
        /// <summary>
        /// Validation type
        /// </summary>
        ExcelDataValidationType ValidationType { get; }
        /// <summary>
        /// Controls how Excel will handle invalid values.
        /// </summary>
        ExcelDataValidationWarningStyle ErrorStyle{ get; set; }
        /// <summary>
        /// True if input message should be shown
        /// </summary>
        bool? AllowBlank { get; set; }
        /// <summary>
        /// True if input message should be shown
        /// </summary>
        bool? ShowInputMessage { get; set; }
        /// <summary>
        /// True if error message should be shown.
        /// </summary>
        bool? ShowErrorMessage { get; set; }
        /// <summary>
        /// Title of error message box (see property ShowErrorMessage)
        /// </summary>
        string ErrorTitle { get; set; }
        /// <summary>
        /// Error message box text (see property ShowErrorMessage)
        /// </summary>
        string Error { get; set; }
        /// <summary>
        /// Title of info box if input message should be shown (see property ShowInputMessage)
        /// </summary>
        string PromptTitle { get; set; }
        /// <summary>
        /// Info message text (see property ShowErrorMessage)
        /// </summary>
        string Prompt { get; set; }
        /// <summary>
        /// True if the current validation type allows operator.
        /// </summary>
        bool AllowsOperator { get; }
        /// <summary>
        /// Validates the state of the validation.
        /// </summary>
        void Validate();


    }
}
