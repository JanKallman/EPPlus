/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * EPPlus is a fork of the ExcelPackage project
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
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2010-03-14
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Security.Cryptography;

namespace OfficeOpenXml
{
    /// <summary>
    /// Sheet protection
    /// </summary>
    public class ExcelSheetProtection : XmlHelper
    {
        public ExcelSheetProtection (XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {

        }        
        private const string _isProtectedPath="d:sheetProtection/@sheet";
        public bool IsProtected
        {
            get
            {
                return GetXmlNodeBool(_isProtectedPath, false);
            }
            set
            {
                SetXmlNodeBool(_isProtectedPath, value, false);
                if (value)
                {
                    AllowObject = true;
                    AllowScenarios = true;
                }
                else
                {
                    DeleteNode(_isProtectedPath); //delete the whole sheetprotection node
                }
            }
        }
        private const string _allowObjectPath="d:sheetProtection/@objects";
        public bool AllowObject
        {
            get
            {
                return GetXmlNodeBool(_allowObjectPath, false);
            }
            set
            {
                SetXmlNodeBool(_allowObjectPath, value, false);
            }
        }
        private const string _allowScenariosPath="d:sheetProtection/@scenarios";
        public bool AllowScenarios
        {
            get
            {
                return GetXmlNodeBool(_allowScenariosPath, false);
            }
            set
            {
                SetXmlNodeBool(_allowScenariosPath, value, false);
            }
        }
        private const string _allowFormatCellsPath="d:sheetProtection/@formatCells";
        public bool AllowFormatCells 
        {
            get
            {
                return GetXmlNodeBool(_allowFormatCellsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowFormatCellsPath, value, true );
            }
        }
        private const string _allowFormatColumnsPath = "d:sheetProtection/@formatColumns";
        public bool AllowFormatColumns
        {
            get
            {
                return GetXmlNodeBool(_allowFormatColumnsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowFormatColumnsPath, value, true);
            }
        }
        private const string _allowFormatRowsPath = "d:sheetProtection/@formatRows";
        public bool AllowFormatRows
        {
            get
            {
                return GetXmlNodeBool(_allowFormatRowsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowFormatRowsPath, value, true);
            }
        }

        private const string _allowInsertColumnsPath = "d:sheetProtection/@insertColumns ";
        public bool AllowInsertColumns
        {
            get
            {
                return GetXmlNodeBool(_allowInsertColumnsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowInsertColumnsPath, value, true);
            }
        }

        private const string _allowInsertRowsPath = "d:sheetProtection/@insertRows";
        public bool AllowInsertRows
        {
            get
            {
                return GetXmlNodeBool(_allowInsertRowsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowInsertRowsPath, value, true);
            }
        }
        private const string _allowInsertHyperlinksPath = "d:sheetProtection/@insertHyperlinks";
        public bool AllowInsertHyperlinks
        {
            get
            {
                return GetXmlNodeBool(_allowInsertHyperlinksPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowInsertHyperlinksPath, value, true);
            }
        }
        private const string _allowDeleteColumns = "d:sheetProtection/@deleteColumns";
        public bool AllowDeleteColumns
        {
            get
            {
                return GetXmlNodeBool(_allowDeleteColumns, true);
            }
            set
            {
                SetXmlNodeBool(_allowDeleteColumns, value, true);
            }
        }
        private const string _allowDeleteRowsPath = "d:sheetProtection/@deleteRows";
        public bool AllowDeleteRows
        {
            get
            {
                return GetXmlNodeBool(_allowDeleteRowsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowDeleteRowsPath, value, true);
            }
        }

        private const string _allowSortPath = "d:sheetProtection/@sort";
        public bool AllowSort
        {
            get
            {
                return GetXmlNodeBool(_allowSortPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowSortPath, value, true);
            }
        }

        private const string _allowAutoFilterPath = "d:sheetProtection/@autoFilter";
        public bool AllowAutoFilter
        {
            get
            {
                return GetXmlNodeBool(_allowAutoFilterPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowAutoFilterPath, value, true);
            }
        }
        private const string _allowPivotTablesPath = "d:sheetProtection/@pivotTables";
        public bool AllowPivotTables
        {
            get
            {
                return GetXmlNodeBool(_allowPivotTablesPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowPivotTablesPath, value, true);
            }
        }

        private const string _passwordPath = "d:sheetProtection/@password";
        /// <summary>
        /// Sets a password for the sheet.
        /// </summary>
        /// <param name="Password"></param>
        public void SetPassword(string Password)
        {
            if (IsProtected == false) IsProtected = true;

            Password = Password.Trim();
            if (Password == "")
            {
                var node = TopNode.SelectSingleNode(_passwordPath, NameSpaceManager);
                if (node != null)
                {
                    (node as XmlAttribute).OwnerElement.Attributes.Remove(node as XmlAttribute);
                }
                return;
            }

            //Calculate the hash
            //Thanks to Kohei Yoshida for the sample http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/
            ushort hash = 0;                
            for (int i = Password.Length-1; i >= 0; i--) 
            {
                hash ^= Password[i];
                hash = (ushort)(((ushort)((hash >> 14) & 0x01))
                                |
                                ((ushort)((hash << 1) & 0x7FFF)));
            }

            hash ^= (0x8000 | ('N' << 8) | 'K');
            hash ^= (ushort)Password.Length;

            SetXmlNode(_passwordPath, ((int)hash).ToString("x"));
        }

    }
}
