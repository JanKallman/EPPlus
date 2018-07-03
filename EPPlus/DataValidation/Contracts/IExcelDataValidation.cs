﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2011-01-08
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
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
