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
using System.Text.RegularExpressions;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Excel datavalidation
    /// </summary>
    public class ExcelDataValidation : XmlHelper
    {
        private const string _rootElement = "d:dataValidation";

        private static string CreatePathFromRoot(string path)
        {
            return string.Concat(_rootElement, "/", path);
        }

        private readonly string _errorStylePath = CreatePathFromRoot("@errorStyle");
        private readonly string _operatorPath = CreatePathFromRoot("@operator");
        private readonly string _showErrorMessagePath = CreatePathFromRoot("@showErrorMessage");
        private readonly string _showInfoMessagePath = CreatePathFromRoot("@showInputMessage");
        private readonly string _typeMessagePath = CreatePathFromRoot("@type");
        private readonly string _sqrefPath = CreatePathFromRoot("@sqref");
        private readonly string _allowBlankPath = CreatePathFromRoot("@allowBlank");
        private readonly string _formula1Path = CreatePathFromRoot("d:formula1");
        private readonly string _formula2Path = CreatePathFromRoot("d:formula2");

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="namespaceManager">Xml namespace manager</param>
        /// <param name="topNode">Xml top node</param>
        /// <param name="validationType">Data validation type</param>
        /// <param name="address">address for data validation</param>
        internal ExcelDataValidation(XmlNamespaceManager namespaceManager, XmlNode topNode, string address, ExcelDataValidationType validationType)
            : base(namespaceManager, topNode)
        {
            Require.Argument(address).IsNotNullOrEmpty("address");
            Require.Argument(topNode).IsNotNull("topNode");
            ValidationType = validationType;
            _address = new ExcelAddress(address);
            _topNode = topNode;
            Init();
        }

        private ExcelAddress _address;
        private XmlNode _topNode;

        private void Init()
        {
            // set schema node order
            SchemaNodeOrder = new string[]{ 
                "type", 
                "errorStyle", 
                "operator", 
                "allowBlank",
                "showInputMessage", 
                "showErrorMessage", 
                "errorTitle", 
                "error", 
                "promptTitle", 
                "prompt", 
                "sqref",
                "formula1",
                "formula2"
            };
        }

        private void SetNullableBoolValue(string path, bool? val)
        {
            if (val.HasValue)
            {
                SetXmlNodeBool(path, val.Value);
            }
        }

        internal void SaveToXml()
        {
            CreateNode(_rootElement);
            SetXmlNodeString(_typeMessagePath, ValidationType.SchemaName);
            if (ValidationType.AllowOperator)
            {
                SetXmlNodeString(_operatorPath, Operator.ToString());
            }
            if (ErrorStyle != ExcelDataValidationWarningStyle.undefined)
            {
                SetXmlNodeString(_errorStylePath, ErrorStyle.ToString());
            }
            SetNullableBoolValue(_showInfoMessagePath, ShowInputMessage);
            SetNullableBoolValue(_showErrorMessagePath, ShowErrorMessage);
            SetNullableBoolValue(_allowBlankPath, AllowBlank);
            // validate and set Formula1
            if (string.IsNullOrEmpty(Formula1))
            {
                throw new InvalidOperationException("Formula1 cannot be empty");
            }
            // lists should always be comma separated
            else if (ValidationType.ValidationType == eDataValidationType.List)
            {
                if (!Formula1.Contains(','))
                {
                    throw new FormatException("When validationtype is list, Formula 1 should be a commaseparated list of values");
                }
            }
            SetXmlNodeString(_formula1Path, Formula1);
            if (Operator == ExcelDataValidationOperator.between || Operator == ExcelDataValidationOperator.notBetween)
            {
                if (string.IsNullOrEmpty(Formula2))
                {
                    throw new InvalidOperationException("Formula2 must be set if operator is 'between' or 'notBetween'");
                }
                SetXmlNodeString(_formula2Path, Formula2);
            }
            SetXmlNodeString(_sqrefPath, _address.Address);
        }

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
        /// Operator
        /// </summary>
        public ExcelDataValidationOperator Operator
        {
            get;
            set;
        }

        /// <summary>
        /// Warning style
        /// </summary>
        public ExcelDataValidationWarningStyle ErrorStyle
        {
            set;
            get;
        }

        /// <summary>
        /// True if blanks should be allowed
        /// </summary>
        public bool? AllowBlank
        {
            get;
            set;
        }

        /// <summary>
        /// True if input message should be shown
        /// </summary>
        public bool? ShowInputMessage
        {
            get;
            set;
        }

        /// <summary>
        /// True if error message should be shown
        /// </summary>
        public bool? ShowErrorMessage
        {
            set;
            get;
        }

        /// <summary>
        /// Formula 1
        /// </summary>
        public string Formula1
        {
            set;
            get;
        }

        /// <summary>
        /// Formula 2
        /// </summary>
        public string Formula2
        {
            get;
            set;
        }

        #endregion
    }
}
