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
    public sealed class ExcelSheetProtection : XmlHelper
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
                    AllowEditObject = true;
                    AllowEditScenarios = true;
                }
                else
                {
                    DeleteAllNode(_isProtectedPath); //delete the whole sheetprotection node
                }
            }
        }
        private const string _allowSelectLockedCellsPath = "d:sheetProtection/@selectLockedCells";
        /// <summary>
        /// Allow users to select locked cells
        /// </summary>
        public bool AllowSelectLockedCells
        {
            get
            {
                return !GetXmlNodeBool(_allowSelectLockedCellsPath, false);
            }
            set
            {
                SetXmlNodeBool(_allowSelectLockedCellsPath, !value, false);
            }
        }
        private const string _allowSelectUnlockedCellsPath = "d:sheetProtection/@selectUnlockedCells";
        /// <summary>
        /// Allow users to select unlocked cells
        /// </summary>
        public bool AllowSelectUnlockedCells
        {
            get
            {
                return !GetXmlNodeBool(_allowSelectUnlockedCellsPath, false);
            }
            set
            {
                SetXmlNodeBool(_allowSelectUnlockedCellsPath, !value, false);
            }
        }        
        private const string _allowObjectPath="d:sheetProtection/@objects";
        /// <summary>
        /// Allow users to edit objects
        /// </summary>
        public bool AllowEditObject
        {
            get
            {
                return !GetXmlNodeBool(_allowObjectPath, false);
            }
            set
            {
                SetXmlNodeBool(_allowObjectPath, !value, false);
            }
        }
        private const string _allowScenariosPath="d:sheetProtection/@scenarios";
        /// <summary>
        /// Allow users to edit senarios
        /// </summary>
        public bool AllowEditScenarios
        {
            get
            {
                return !GetXmlNodeBool(_allowScenariosPath, false);
            }
            set
            {
                SetXmlNodeBool(_allowScenariosPath, !value, false);
            }
        }
        private const string _allowFormatCellsPath="d:sheetProtection/@formatCells";
        /// <summary>
        /// Allow users to format cells
        /// </summary>
        public bool AllowFormatCells 
        {
            get
            {
                return !GetXmlNodeBool(_allowFormatCellsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowFormatCellsPath, !value, true );
            }
        }
        private const string _allowFormatColumnsPath = "d:sheetProtection/@formatColumns";
        /// <summary>
        /// Allow users to Format columns
        /// </summary>
        public bool AllowFormatColumns
        {
            get
            {
                return !GetXmlNodeBool(_allowFormatColumnsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowFormatColumnsPath, !value, true);
            }
        }
        private const string _allowFormatRowsPath = "d:sheetProtection/@formatRows";
        /// <summary>
        /// Allow users to Format rows
        /// </summary>
        public bool AllowFormatRows
        {
            get
            {
                return !GetXmlNodeBool(_allowFormatRowsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowFormatRowsPath, !value, true);
            }
        }

        private const string _allowInsertColumnsPath = "d:sheetProtection/@insertColumns ";
        /// <summary>
        /// Allow users to insert columns
        /// </summary>
        public bool AllowInsertColumns
        {
            get
            {
                return !GetXmlNodeBool(_allowInsertColumnsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowInsertColumnsPath, !value, true);
            }
        }

        private const string _allowInsertRowsPath = "d:sheetProtection/@insertRows";
        /// <summary>
        /// Allow users to Format rows
        /// </summary>
        public bool AllowInsertRows
        {
            get
            {
                return !GetXmlNodeBool(_allowInsertRowsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowInsertRowsPath, !value, true);
            }
        }
        private const string _allowInsertHyperlinksPath = "d:sheetProtection/@insertHyperlinks";
        /// <summary>
        /// Allow users to insert hyperlinks
        /// </summary>
        public bool AllowInsertHyperlinks
        {
            get
            {
                return !GetXmlNodeBool(_allowInsertHyperlinksPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowInsertHyperlinksPath, !value, true);
            }
        }
        private const string _allowDeleteColumns = "d:sheetProtection/@deleteColumns";
        /// <summary>
        /// Allow users to delete columns
        /// </summary>
        public bool AllowDeleteColumns
        {
            get
            {
                return !GetXmlNodeBool(_allowDeleteColumns, true);
            }
            set
            {
                SetXmlNodeBool(_allowDeleteColumns, !value, true);
            }
        }
        private const string _allowDeleteRowsPath = "d:sheetProtection/@deleteRows";
        /// <summary>
        /// Allow users to delete rows
        /// </summary>
        public bool AllowDeleteRows
        {
            get
            {
                return !GetXmlNodeBool(_allowDeleteRowsPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowDeleteRowsPath, !value, true);
            }
        }

        private const string _allowSortPath = "d:sheetProtection/@sort";
        /// <summary>
        /// Allow users to sort a range
        /// </summary>
        public bool AllowSort
        {
            get
            {
                return !GetXmlNodeBool(_allowSortPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowSortPath, !value, true);
            }
        }

        private const string _allowAutoFilterPath = "d:sheetProtection/@autoFilter";
        public bool AllowAutoFilter
        {
            get
            {
                return !GetXmlNodeBool(_allowAutoFilterPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowAutoFilterPath, !value, true);
            }
        }
        private const string _allowPivotTablesPath = "d:sheetProtection/@pivotTables";
        public bool AllowPivotTables
        {
            get
            {
                return !GetXmlNodeBool(_allowPivotTablesPath, true);
            }
            set
            {
                SetXmlNodeBool(_allowPivotTablesPath, !value, true);
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

            SetXmlNodeString(_passwordPath, ((int)hash).ToString("x"));
        }

    }
}
