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
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A field Item. Used for grouping
    /// </summary>
    public class ExcelPivotTableFieldItem : XmlHelper
    {
        ExcelPivotTableField _field;
        internal ExcelPivotTableFieldItem(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTableField field) :
            base(ns, topNode)
        {
           _field = field;
        }
        /// <summary>
        /// The text. Unique values only
        /// </summary>
        public string Text
        {
            get
            {
                return GetXmlNodeString("@n");
            }
            set
            {
                if(string.IsNullOrEmpty(value))
                {
                    DeleteNode("@n");
                    return;
                }
                foreach (var item in _field.Items)
                {
                    if (item.Text == value)
                    {
                        throw(new ArgumentException("Duplicate Text"));
                    }
                }
                SetXmlNodeString("@n", value);
            }
        }
        internal int X
        {
            get
            {
                return GetXmlNodeInt("@x"); 
            }
        }
        internal string T
        {
            get
            {
                return GetXmlNodeString("@t");
            }
        }
    }
}
