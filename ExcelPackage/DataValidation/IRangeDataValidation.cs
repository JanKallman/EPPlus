using System;
using System.Collections.Generic;
/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 *  Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Mats Alm   		                Added       		        2011-03-23
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
