/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
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
 * Mats Alm   		                Added       		        2011-03-23
 * Jan Källman		                License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Contracts;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Provides functionality for adding datavalidation to a range (<see cref="ExcelRangeBase"/>). Each method will
    /// return a configurable validation.
    /// </summary>
    public interface IRangeDataValidation
    {
        /// <summary>
        /// Adds a <see cref="IExcelDataValidationInt"/> to the range
        /// </summary>
        /// <returns>A <see cref="ExcelDataValidationInt"/> that can be configured for integer data validation</returns>
        IExcelDataValidationInt AddIntegerDataValidation();
        /// <summary>
        /// Adds a <see cref="ExcelDataValidationDecimal"/> to the range
        /// </summary>
        /// <returns>A <see cref="ExcelDataValidationDecimal"/> that can be configured for decimal data validation</returns>
        IExcelDataValidationDecimal AddDecimalDataValidation();
        /// <summary>
        /// Adds a <see cref="ExcelDataValidationDateTime"/> to the range
        /// </summary>
        /// <returns>A <see cref="ExcelDataValidationDecimal"/> that can be configured for datetime data validation</returns>
        IExcelDataValidationDateTime AddDateTimeDataValidation();
        /// <summary>
        /// Adds a <see cref="IExcelDataValidationList"/> to the range
        /// </summary>
        /// <returns>A <see cref="ExcelDataValidationList"/> that can be configured for datetime data validation</returns>
        IExcelDataValidationList AddListDataValidation();
        /// <summary>
        /// Adds a <see cref="IExcelDataValidationInt"/> regarding text length validation to the range.
        /// </summary>
        /// <returns></returns>
        IExcelDataValidationInt AddTextLengthDataValidation();
        /// <summary>
        /// Adds a <see cref="IExcelDataValidationTime"/> to the range.
        /// </summary>
        /// <returns>A <see cref="IExcelDataValidationTime"/> that can be configured for time data validation</returns>
        IExcelDataValidationTime AddTimeDataValidation();
        /// <summary>
        /// Adds a <see cref="IExcelDataValidationCustom"/> to the range.
        /// </summary>
        /// <returns>A <see cref="IExcelDataValidationCustom"/> that can be configured for custom validation</returns>
        IExcelDataValidationCustom AddCustomDataValidation();
    }
}
