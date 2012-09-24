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
 *******************************************************************************
 * Jan Källman		Added		12-APR-2012
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;

namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// Vba security properties
    /// </summary>
    public class ExcelVbaProtection
    {
        ExcelVbaProject _project;
        internal ExcelVbaProtection(ExcelVbaProject project)
        {
            _project = project;
            VisibilityState = true;
        }
        /// <summary>
        /// Specifies whether access to the VBA project was restricted by the user
        /// </summary>
        public bool UserProtected { get; internal set; }
        /// <summary>
        /// Specifies whether access to the VBA project was restricted by the VBA host application
        /// </summary>
        public bool HostProtected { get; internal set; }
        /// <summary>
        /// Specifies whether access to the VBA project was restricted by the VBA project editor
        /// </summary>
        public bool VbeProtected { get; internal set; }
        /// <summary>
        /// Specifies whether the VBA project is visible.
        /// </summary>
        public bool VisibilityState { get; internal set; }
        internal byte[] PasswordHash { get; set; }
        internal byte[] PasswordKey { get; set; }
        /// <summary>
        /// Password protect the VBA project.
        /// An empty string or null will remove the password protection
        /// </summary>
        /// <param name="Password">The password</param>
        public void SetPassword(string Password)
        {

            if (string.IsNullOrEmpty(Password))
            {
                PasswordHash = null;
                PasswordKey = null;
                VbeProtected = false;
                HostProtected = false;
                UserProtected = false;
                VisibilityState = true;
                _project.ProjectID = "{5DD90D76-4904-47A2-AF0D-D69B4673604E}";
            }
            else
            {
                //Join Password and Key
                byte[] data;
                //Set the key
                PasswordKey = new byte[4];
                RandomNumberGenerator r = RandomNumberGenerator.Create();
                r.GetBytes(PasswordKey);

                data = new byte[Password.Length + 4];
                Array.Copy(Encoding.GetEncoding(_project.CodePage).GetBytes(Password), data, Password.Length);
                VbeProtected = true;
                VisibilityState = false;
                Array.Copy(PasswordKey, 0, data, data.Length - 4, 4);

                //Calculate Hash
                var provider = SHA1.Create();
                PasswordHash = provider.ComputeHash(data);
                _project.ProjectID = "{00000000-0000-0000-0000-000000000000}";
            }
        }
        //public void ValidatePassword(string Password)                     
        //{

        //}        
    }
}
