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
using OfficeOpenXml.Utils;
using System.Xml;

namespace OfficeOpenXml
{
    /// <summary>
    /// Excel datavalidation
    /// </summary>
    public class ExcelDataValidation : XmlHelper
    {

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="namespaceManager">Xml namespace manager</param>
        /// <param name="topNode">Xml top node</param>
        /// <param name="validationType">Data validation type</param>
        /// <param name="address">address for data validation</param>
        /// <param name="worksheet">worksheet</param>
        internal ExcelDataValidation(XmlNamespaceManager namespaceManager, XmlNode topNode,  ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : base(namespaceManager, topNode)
        {
            Require.Argument(worksheet).IsNotNull("worksheet");
            Require.Argument(address).IsNotNullOrEmpty(address);
            ValidationType = validationType;
            _range = new ExcelRange(worksheet, address);
            _worksheet = worksheet;
        }

        private ExcelWorksheet _worksheet;
        private ExcelRange _range;

        #region Public properties
        /// <summary>
        /// Validation type
        /// </summary>
        public ExcelDataValidationType ValidationType
        {
            get;
            private set;
        }

        /// <summary>
        /// True if blanks should be allowed
        /// </summary>
        public bool AllowBlanks
        {
            get;
            set;
        }

        /// <summary>
        /// True if input message should be shown
        /// </summary>
        public bool ShowInputMessage
        {
            get;
            set;
        }

        /// <summary>
        /// True if error message should be shown
        /// </summary>
        public bool ShowErrorMessage
        {
            set;
            get;
        }

        #endregion
    }
}
