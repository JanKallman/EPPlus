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
 * Jan Källman		    Initial Release		        2010-03-14
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Security.Cryptography;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml
{
    /// <summary>
    /// Sheet protection
    ///<seealso cref="ExcelEncryption"/> 
    ///<seealso cref="ExcelProtection"/> 
    /// </summary>
    public sealed class ExcelSheetProtection : XmlHelper
    {
        internal ExcelSheetProtection (XmlNamespaceManager nsm, XmlNode topNode,ExcelWorksheet ws) :
            base(nsm, topNode)
        {
            SchemaNodeOrder = ws.SchemaNodeOrder;
        }        
        private const string _isProtectedPath="d:sheetProtection/@sheet";
        /// <summary>
        /// If the worksheet is protected.
        /// </summary>
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

        private const string _allowInsertColumnsPath = "d:sheetProtection/@insertColumns";
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
        /// <summary>
        /// Allow users to use autofilters
        /// </summary>
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
        /// <summary>
        /// Allow users to use pivottables
        /// </summary>
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

            int hash = EncryptedPackageHandler.CalculatePasswordHash(Password);
            SetXmlNodeString(_passwordPath, ((int)hash).ToString("x"));
        }

    }
}
