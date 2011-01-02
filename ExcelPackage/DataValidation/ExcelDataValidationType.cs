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
 * Mats Alm   		                Added       		        2011-01-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Enum for available data validation types
    /// </summary>
    public enum eDataValidationType
    {
        /// <summary>
        /// Integer value
        /// </summary>
        Whole,
        /// <summary>
        /// Decimal values
        /// </summary>
        Decimal,
        /// <summary>
        /// List of values
        /// </summary>
        List
    }

    /// <summary>
    /// Types of datavalidation
    /// </summary>
    public class ExcelDataValidationType
    {
        private ExcelDataValidationType(eDataValidationType validationType, bool allowOperator, string schemaName)
        {
            ValidationType = validationType;
            AllowOperator = allowOperator;
            SchemaName = schemaName;
        }

        /// <summary>
        /// Validation type
        /// </summary>
        public eDataValidationType ValidationType
        {
            get;
            private set;
        }

        internal string SchemaName
        {
            get;
            private set;
        }

        /// <summary>
        /// This type allows operator to be set
        /// </summary>
        internal bool AllowOperator
        {

            get;
            private set;
        }

        /// <summary>
        /// Returns a validation type by <see cref="eDataValidationType"/>
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        internal static ExcelDataValidationType GetByValidationType(eDataValidationType type)
        {
            switch (type)
            {
                case eDataValidationType.Whole:
                    return ExcelDataValidationType.Whole;
                case eDataValidationType.List:
                    return ExcelDataValidationType.List;
                case eDataValidationType.Decimal:
                    return ExcelDataValidationType.Decimal;
                default:
                    throw new InvalidOperationException("Non supported Validationtype : " + type.ToString());
            }
        }

        /// <summary>
        /// Overridden Equals, compares on internal validation type
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            if (!(obj is ExcelDataValidationType))
            {
                return false;
            }
            return ((ExcelDataValidationType)obj).ValidationType == ValidationType;
        }

        /// <summary>
        /// Overrides GetHashCode()
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        /// <summary>
        /// Integer values
        /// </summary>
        private static ExcelDataValidationType _whole;
        public static ExcelDataValidationType Whole
        {
            get 
            {
                if(_whole == null)
                {
                    _whole = new ExcelDataValidationType(eDataValidationType.Whole, true, "whole"); 
                }
                return _whole;
            }
        }

        /// <summary>
        /// List of allowed values
        /// </summary>
        private static ExcelDataValidationType _list;
        public static ExcelDataValidationType List
        {
            get
            {
                if (_list == null)
                {
                    _list = new ExcelDataValidationType(eDataValidationType.List, false, "list");
                }
                return _list;
            }
        }

        private static ExcelDataValidationType _decimal;
        public static ExcelDataValidationType Decimal
        {
            get
            {
                if (_decimal == null)
                {
                    _decimal = new ExcelDataValidationType(eDataValidationType.Decimal, true, "decimal");
                }
                return _decimal;
            }
        }
    }
}
