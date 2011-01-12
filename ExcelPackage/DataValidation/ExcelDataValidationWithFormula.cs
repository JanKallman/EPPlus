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
 * Mats Alm   		                Added       		        2011-01-08
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using System.Xml;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// A validation containing a formula
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelDataValidationWithFormula<T> : ExcelDataValidation
        where T : IExcelDataValidationFormula
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="validationType"></param>
        internal ExcelDataValidationWithFormula(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : this(worksheet, address, validationType, null)
        {

        }

         /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet">Worksheet that owns the validation</param>
        /// <param name="itemElementNode">Xml top node (dataValidations)</param>
        /// <param name="validationType">Data validation type</param>
        /// <param name="address">address for data validation</param>
        internal ExcelDataValidationWithFormula(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet, address, validationType, itemElementNode)
        {
            
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="worksheet">Worksheet that owns the validation</param>
        /// <param name="itemElementNode">Xml top node (dataValidations)</param>
        /// <param name="validationType">Data validation type</param>
        /// <param name="address">address for data validation</param>
        /// <param name="namespaceManager">for test purposes</param>
        internal ExcelDataValidationWithFormula(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode, XmlNamespaceManager namespaceManager)
            : base(worksheet, address, validationType, itemElementNode, namespaceManager)
        {

        }


        /// <summary>
        /// Formula - Either a {T} value (except for custom validation) or a spreadsheet formula
        /// </summary>
        public T Formula
        {
            get;
            protected set;
        }

        public override void Validate()
        {
            base.Validate();
            if (Operator == ExcelDataValidationOperator.between || Operator == ExcelDataValidationOperator.notBetween)
            {
                if (string.IsNullOrEmpty(Formula2Internal))
                {
                    throw new InvalidOperationException("Validation of " + Address.Address + " failed: Formula2 must be set if operator is 'between' or 'notBetween'");
                }
            }
        }
    }
}
