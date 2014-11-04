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
 * Jan Källman		Added		21-MAR-2011
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A page / report filter field
    /// </summary>
    public class ExcelPivotTablePageFieldSettings  : XmlHelper
    {
        ExcelPivotTableField _field;
        internal ExcelPivotTablePageFieldSettings(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field, int index) :
            base(ns, topNode)
        {
            if (GetXmlNodeString("@hier")=="")
            {
                Hier = -1;
            }
            _field = field;
        }
        internal int Index 
        { 
            get
            {
                return GetXmlNodeInt("@fld");
            }
            set
            {
                SetXmlNodeString("@fld",value.ToString());
            }
        }
        /// <summary>
        /// The Name of the field
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value);
            }
        }
        /***** Dont work. Need items to be populated. ****/
        ///// <summary>
        ///// The selected item 
        ///// </summary>
        //public int SelectedItem
        //{
        //    get
        //    {
        //        return GetXmlNodeInt("@item");
        //    }
        //    set
        //    {
        //        if (value < 0) throw new InvalidOperationException("Can't be negative");
        //        SetXmlNodeString("@item", value.ToString());
        //    }
        //}
        internal int NumFmtId
        {
            get
            {
                return GetXmlNodeInt("@numFmtId");
            }
            set
            {
                SetXmlNodeString("@numFmtId", value.ToString());
            }
        }
        internal int Hier
        {
            get
            {
                return GetXmlNodeInt("@hier");
            }
            set
            {
                SetXmlNodeString("@hier", value.ToString());
            }
        }
    }
}
