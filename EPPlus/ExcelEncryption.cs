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
 ***************************************************************************************
 * This class is created with the help of the MS-OFFCRYPTO PDF documentation... http://msdn.microsoft.com/en-us/library/cc313071(office.12).aspx
 * Decrypytion library for Office Open XML files(Lyquidity) and Sminks very nice example 
 * on "Reading compound documents in c#" on Stackoverflow. Many thanks!
 ***************************************************************************************
 *  
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		10-AUG-2010
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// Encryption Algorithm
    /// </summary>
    public enum EncryptionAlgorithm
    {
        /// <summary>
        /// 128-bit AES. Default
        /// </summary>
        AES128,
        /// <summary>
        /// 192-bit AES.
        /// </summary>
        AES192,
        /// <summary>
        /// 256-bit AES. 
        /// </summary>
        AES256
    }
    /// <summary>
    /// The major version of the Encryption 
    /// </summary>
    public enum EncryptionVersion
    {
        /// <summary>
        /// Standard Encryption.
        /// Used in Excel 2007
        /// </summary>
        Version3,
        /// <summary>
        /// Agile Encryption.
        /// Used in Excel 2010-
        /// Default.
        /// </summary>
        Version4
    }
    /// <summary>
    /// How and if the workbook is encrypted
    ///<seealso cref="ExcelProtection"/> 
    ///<seealso cref="ExcelSheetProtection"/> 
    /// </summary>
    public class ExcelEncryption
    {
        /// <summary>
        /// Constructor
        /// </summary>
        internal ExcelEncryption()
        {
            Algorithm = EncryptionAlgorithm.AES128;
        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="encryptionAlgorithm">Algorithm used to encrypt the package. Default is AES128</param>
        internal ExcelEncryption(EncryptionAlgorithm encryptionAlgorithm)
        {
            Algorithm = encryptionAlgorithm;
        }        
        bool _isEncrypted = false;
        /// <summary>
        /// Is the package encrypted
        /// </summary>
        public bool IsEncrypted 
        {
            get
            {
                return _isEncrypted;
            }
            set
            {
                _isEncrypted = value;
                if (_isEncrypted)
                {
                    if (_password == null) _password = "";
                }
                else
                {
                    _password = null;
                }
            }
        }
        string _password=null;
        /// <summary>
        /// The password used to encrypt the workbook.
        /// </summary>
        public string Password 
        {
            get
            {
                return _password;
            }
            set
            {
                _password = value;
                _isEncrypted = (value != null);
            }
        }
        /// <summary>
        /// Algorithm used for encrypting the package. Default is AES 128-bit
        /// </summary>
        public EncryptionAlgorithm Algorithm { get; set; }
        public EncryptionVersion Version
        {
            get;
            set;
        }
    }
}
