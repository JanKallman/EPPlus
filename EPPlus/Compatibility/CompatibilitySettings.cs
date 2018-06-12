/*******************************************************************************
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
 * Jan Källman		    Added       		        2017-11-02
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;

namespace OfficeOpenXml.Compatibility
{
    /// <summary>
    /// Settings to stay compatible with older versions of EPPlus
    /// </summary>
    public class CompatibilitySettings
    {
        private ExcelPackage excelPackage;


        internal CompatibilitySettings(ExcelPackage excelPackage)
        {
            this.excelPackage = excelPackage;
        }
#if Core
        /// <summary>
        /// If the worksheets collection of the ExcelWorkbook class is 1 based.
        /// This property can be set from appsettings.json file.
        /// <code>
        ///     {
        ///       "EPPlus": {
        ///         "ExcelPackage": {
        ///           "Compatibility": {
        ///             "IsWorksheets1Based": false //Default value is false
        ///           }
        ///         }
        ///       }
        ///     }
        /// </code>
        /// </summary>
#else
        /// <summary>
        /// If the worksheets collection of the ExcelWorkbook class is 1 based.
        /// This property can be set from app.config file.
        /// <code>
        ///   <appSettings>
        ///    <!--Set worksheets collection to start from zero.Default is 1, for backward compatibility reasons -->  
        ///    <add key = "EPPlus:ExcelPackage.Compatibility.IsWorksheets1Based" value="false" />
        ///   </appSettings>
        /// </code>
        /// </summary>
#endif

        public bool IsWorksheets1Based
        {
            get
            {
                return excelPackage._worksheetAdd==1;
            }
            set
            {
                excelPackage._worksheetAdd = value ? 1 : 0;
                if(excelPackage._workbook!=null && excelPackage._workbook._worksheets!=null)
                {
                    excelPackage.Workbook.Worksheets.ReindexWorksheetDictionary();

                }
            }
        }
    }
}
