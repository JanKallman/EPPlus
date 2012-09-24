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
 * Jan Källman		Added		26-MAR-2012
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.VBA
{
    /// <summary>
    /// A VBA reference
    /// </summary>
    public class ExcelVbaReference
    {
        public ExcelVbaReference()
        {
            ReferenceRecordID = 0xD;
        }
        /// <summary>
        /// The reference record ID. See MS-OVBA documentation for more info. 
        /// </summary>
        public int ReferenceRecordID { get; internal set; }
        /// <summary>
        /// The name of the reference
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// LibID
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string Libid { get; set; }
        /// <summary>
        /// A string representation of the object (the Name)
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Name;
        }
    }
    /// <summary>
    /// A reference to a twiddled type library
    /// </summary>
    public class ExcelVbaReferenceControl : ExcelVbaReference
    {
        public ExcelVbaReferenceControl()
        {
            ReferenceRecordID = 0x2F;
        }
        /// <summary>
        /// LibIdExternal 
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string LibIdExternal { get; set; }
        /// <summary>
        /// LibIdTwiddled
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string LibIdTwiddled { get; set; }
        /// <summary>
        /// A GUID that specifies the Automation type library the extended type library was generated from.
        /// </summary>
        public Guid OriginalTypeLib { get; set; }
        internal uint Cookie { get; set; }
    }
    /// <summary>
    /// A reference to an external VBA project
    /// </summary>
    public class ExcelVbaReferenceProject : ExcelVbaReference
    {
        public ExcelVbaReferenceProject()
        {
            ReferenceRecordID = 0x0E;
        }
        /// <summary>
        /// LibIdRelative
        /// For more info check MS-OVBA 2.1.1.8 LibidReference and 2.3.4.2.2 PROJECTREFERENCES
        /// </summary>
        public string LibIdRelative { get; set; }
        /// <summary>
        /// Major version of the referenced VBA project
        /// </summary>
        public uint MajorVersion { get; set; }
        /// <summary>
        /// Minor version of the referenced VBA project
        /// </summary>
        public ushort MinorVersion { get; set; }
    }
}
