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
 * Jan Källman		Added		2009-12-30
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Chart surface settings
    /// </summary>
    public class ExcelChartSurface : XmlHelper
    {
        internal ExcelChartSurface(XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           SchemaNodeOrder = new string[] { "thickness", "spPr", "pictureOptions" };
       }
       #region "Public properties"
        const string THICKNESS_PATH = "c:thickness/@val";
       /// <summary>
       /// Show the values 
       /// </summary>
        public int Thickness
       {
           get
           {
               return GetXmlNodeInt(THICKNESS_PATH);
           }
           set
           {
               if(value < 0 && value > 9)
               {
                   throw (new ArgumentOutOfRangeException("Thickness out of range. (0-9)"));
               }
               SetXmlNodeString(THICKNESS_PATH, value.ToString());
           }
       }
       ExcelDrawingFill _fill = null;
       /// <summary>
       /// Access fill properties
       /// </summary>
       public ExcelDrawingFill Fill
       {
           get
           {
               if (_fill == null)
               {
                   _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "c:spPr");
               }
               return _fill;
           }
       }
       ExcelDrawingBorder _border = null;
       /// <summary>
       /// Access border properties
       /// </summary>
       public ExcelDrawingBorder Border
       {
           get
           {
               if (_border == null)
               {
                   _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "c:spPr/a:ln");
               }
               return _border;
           }
       }
       #endregion
    }
}
