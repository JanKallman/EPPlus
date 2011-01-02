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
        private const string _itemElementNodeName = "d:dataValidation";

        private static string CreatePathFromRoot(string path)
        {
            return path;
        }

        private readonly string _errorStylePath = CreatePathFromRoot("@errorStyle");
        private readonly string _errorTitlePath = CreatePathFromRoot("@errorTitle");
        private readonly string _errorPath = CreatePathFromRoot("@error");
        private readonly string _promptTitlePath = CreatePathFromRoot("@promptTitle");
        private readonly string _promptPath = CreatePathFromRoot("@prompt");
        private readonly string _operatorPath = CreatePathFromRoot("@operator");
        private readonly string _showErrorMessagePath = CreatePathFromRoot("@showErrorMessage");
        private readonly string _showInputMessagePath = CreatePathFromRoot("@showInputMessage");
        private readonly string _typeMessagePath = CreatePathFromRoot("@type");
        private readonly string _sqrefPath = CreatePathFromRoot("@sqref");
        private readonly string _allowBlankPath = CreatePathFromRoot("@allowBlank");
        protected readonly string _formula1Path = CreatePathFromRoot("d:formula1");
        protected readonly string _formula2Path = CreatePathFromRoot("d:formula2");

        internal ExcelDataValidation(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType)
            : this(worksheet, address, validationType, null)
        { }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="namespaceManager">Xml namespace manager</param>
        /// <param name="topNode">Xml top node (dataValidations)</param>
        /// <param name="validationType">Data validation type</param>
        /// <param name="address">address for data validation</param>
        internal ExcelDataValidation(ExcelWorksheet worksheet, string address, ExcelDataValidationType validationType, XmlNode itemElementNode)
            : base(worksheet.NameSpaceManager)
        {
            Require.Argument(address).IsNotNullOrEmpty("address");
            if (itemElementNode == null)
            {
                TopNode = worksheet.WorksheetXml.SelectSingleNode("//d:dataValidations", worksheet.NameSpaceManager);
                itemElementNode = CreateNode(_itemElementNodeName);
            }
            TopNode = itemElementNode;
            ValidationType = validationType;
            Address = new ExcelAddress(address);
            Init();
        }

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
            else
            {
                DeleteNode(path);
            }
        }

        public void Validate()
        {
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
            if (Operator == ExcelDataValidationOperator.between || Operator == ExcelDataValidationOperator.notBetween)
            {
                if (string.IsNullOrEmpty(Formula2))
                {
                    throw new InvalidOperationException("Formula2 must be set if operator is 'between' or 'notBetween'");
                }
            }
        }

        #region Public properties

        public ExcelAddress Address
        {
            get
            {
                return new ExcelAddress(GetXmlNodeString(_sqrefPath));
            }
            set
            {
                SetXmlNodeString(_sqrefPath, value.Address);
            }
        }
        /// <summary>
        /// Validation type
        /// </summary>
        public ExcelDataValidationType ValidationType
        {
            get
            {
                var typeString = GetXmlNodeString(_typeMessagePath);
                return ExcelDataValidationType.GetBySchemaName(typeString);
            }
            private set
            {
                SetXmlNodeString(_typeMessagePath, value.SchemaName);
            }
        }

        /// <summary>
        /// Operator
        /// </summary>
        public ExcelDataValidationOperator Operator
        {
            get
            {
                var operatorString = GetXmlNodeString(_operatorPath);
                if (!string.IsNullOrEmpty(operatorString))
                {
                    return (ExcelDataValidationOperator)Enum.Parse(typeof(ExcelDataValidationOperator), operatorString);
                }
                return default(ExcelDataValidationOperator);
            }
            set
            {
                if (!ValidationType.AllowOperator)
                {
                    throw new InvalidOperationException("The current validation type does not allow operator to be set");
                }
                SetXmlNodeString(_operatorPath, value.ToString());
            }
        }

        /// <summary>
        /// Warning style
        /// </summary>
        public ExcelDataValidationWarningStyle ErrorStyle
        {
            get
            {
                var errorStyleString = GetXmlNodeString(_errorStylePath);
                if (!string.IsNullOrEmpty(errorStyleString))
                {
                    return (ExcelDataValidationWarningStyle)Enum.Parse(typeof(ExcelDataValidationWarningStyle), errorStyleString);
                }
                return ExcelDataValidationWarningStyle.undefined;
            }
            set
            {
                if(value == ExcelDataValidationWarningStyle.undefined)
                {
                    DeleteNode(_errorStylePath);
                }
                SetXmlNodeString(_errorStylePath, value.ToString());
            }
        }

        /// <summary>
        /// True if blanks should be allowed
        /// </summary>
        public bool? AllowBlank
        {
            get
            {
                return GetXmlNodeBoolNullable(_allowBlankPath);
            }
            set
            {
                SetNullableBoolValue(_allowBlankPath, value);
            }
        }

        /// <summary>
        /// True if input message should be shown
        /// </summary>
        public bool? ShowInputMessage
        {
            get
            {
                return GetXmlNodeBoolNullable(_showInputMessagePath);
            }
            set
            {
                SetNullableBoolValue(_showInputMessagePath, value);
            }
        }

        /// <summary>
        /// True if error message should be shown
        /// </summary>
        public bool? ShowErrorMessage
        {
            get
            {
                return GetXmlNodeBoolNullable(_showErrorMessagePath);
            }
            set
            {
                SetNullableBoolValue(_showErrorMessagePath, value);
            }
        }

        /// <summary>
        /// Title of error message box
        /// </summary>
        public string ErrorTitle
        {
            get
            {
                return GetXmlNodeString(_errorTitlePath);
            }
            set
            {
                SetXmlNodeString(_errorTitlePath, value);
            }
        }

        /// <summary>
        /// Error message box text
        /// </summary>
        public string Error
        {
            get
            {
                return GetXmlNodeString(_errorPath);
            }
            set
            {
                SetXmlNodeString(_errorPath, value);
            }
        }

        public string PromptTitle
        {
            get
            {
                return GetXmlNodeString(_promptTitlePath);
            }
            set
            {
                SetXmlNodeString(_promptTitlePath, value);
            }
        }

        public string Prompt
        {
            get
            {
                return GetXmlNodeString(_promptPath);
            }
            set
            {
                SetXmlNodeString(_promptPath, value);
            }
        }

        /// <summary>
        /// Formula 1
        /// </summary>
        internal virtual string Formula1
        {
            get
            {
                return GetXmlNodeString(_formula1Path);
            }
        }

        /// <summary>
        /// Formula 2
        /// </summary>
        internal virtual string Formula2
        {
            get
            {
                return GetXmlNodeString(_formula2Path);
            }
        }

        #endregion

        protected void SetValue<T>(Nullable<T> val, string path)
            where T : struct
        {
            if (!val.HasValue)
            {
                DeleteNode(path);
            }
            var stringValue = val.Value.ToString().Replace(',', '.');
            SetXmlNodeString(path, stringValue);
        }
    }
}
