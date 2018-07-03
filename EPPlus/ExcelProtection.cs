﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
 * Jan Källman		    Added		10-AUG-2010
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Encryption;
namespace OfficeOpenXml
{
    /// <summary>
    /// Sets protection on the workbook level
    ///<seealso cref="ExcelEncryption"/> 
    ///<seealso cref="ExcelSheetProtection"/> 
    /// </summary>
    public class ExcelProtection : XmlHelper
    {
        internal ExcelProtection(XmlNamespaceManager ns, XmlNode topNode, ExcelWorkbook wb) :
            base(ns, topNode)
        {
            SchemaNodeOrder = wb.SchemaNodeOrder;
        }
        const string workbookPasswordPath = "d:workbookProtection/@workbookPassword";
        /// <summary>
        /// Sets a password for the workbook. This does not encrypt the workbook. 
        /// </summary>
        /// <param name="Password">The password. </param>
        public void SetPassword(string Password)
        {
            if(string.IsNullOrEmpty(Password))
            {
                DeleteNode(workbookPasswordPath);
            }
            else
            {
                SetXmlNodeString(workbookPasswordPath, ((int)EncryptedPackageHandler.CalculatePasswordHash(Password)).ToString("x"));
            }
        }
        const string lockStructurePath = "d:workbookProtection/@lockStructure";
        /// <summary>
        /// Locks the structure,which prevents users from adding or deleting worksheets or from displaying hidden worksheets.
        /// </summary>
        public bool LockStructure
        {
            get
            {
                return GetXmlNodeBool(lockStructurePath, false);
            }
            set
            {
                SetXmlNodeBool(lockStructurePath, value,  false);
            }
        }
        const string lockWindowsPath = "d:workbookProtection/@lockWindows";
        /// <summary>
        /// Locks the position of the workbook window.
        /// </summary>
        public bool LockWindows
        {
            get
            {
                return GetXmlNodeBool(lockWindowsPath, false);
            }
            set
            {
                SetXmlNodeBool(lockWindowsPath, value, false);
            }
        }
        const string lockRevisionPath = "d:workbookProtection/@lockRevision";

        /// <summary>
        /// Lock the workbook for revision
        /// </summary>
        public bool LockRevision
        {
            get
            {
                return GetXmlNodeBool(lockRevisionPath, false);
            }
            set
            {
                SetXmlNodeBool(lockRevisionPath, value, false);
            }
        }
    }
}
