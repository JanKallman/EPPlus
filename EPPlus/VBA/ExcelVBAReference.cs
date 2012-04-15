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
    public class ExcelVbaReference
    {
        public ExcelVbaReference()
        {
            ReferenceRecordID = 0xD;
        }
        public int ReferenceRecordID { get; set; }
        public string Name { get; set; }
        public string Libid { get; set; }
        public override string ToString()
        {
            return Name;
        }
    }
    public class ExcelVbaReferenceControl : ExcelVbaReference
    {
        public ExcelVbaReferenceControl()
        {
            ReferenceRecordID = 0x2F;
        }
        public string LibIdExternal { get; set; }
        public string LibIdTwiddled { get; set; }
        public Guid OriginalTypeLib { get; set; }
        public uint Cookie { get; set; }
    }
    public class ExcelVbaReferenceProject : ExcelVbaReference
    {
        public ExcelVbaReferenceProject()
        {
            ReferenceRecordID = 0x0E;
        }
        public string LibIdRelative { get; set; }
        public uint MajorVersion { get; set; }
        public ushort MinorVersion { get; set; }
    }
}
