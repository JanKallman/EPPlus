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
 * Jan Källman		Added		2009-10-01
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
    /// Datalabel properties
    /// </summary>
    public sealed class ExcelChartSerieDataLabel : ExcelChartDataLabel
    {
       internal ExcelChartSerieDataLabel(XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           CreateNode(positionPath);
           Position = eLabelPosition.Center;
       }

       const string positionPath="c:dLblPos/@val";
       /// <summary>
       /// Position of the labels
       /// </summary>
       public eLabelPosition Position
       {
           get
           {
               return GetPosEnum(GetXmlNodeString(positionPath));
           }
           set
           {
               SetXmlNodeString(positionPath,GetPosText(value));
           }
       }
       ExcelDrawingFill _fill = null;
       /// <summary>
       /// Access fill properties
       /// </summary>
        public new ExcelDrawingFill Fill
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
        public new ExcelDrawingBorder Border
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
       ExcelTextFont _font = null;
       /// <summary>
       /// Access font properties
       /// </summary>
        public new ExcelTextFont Font
       {
           get
           {
               if (_font == null)
               {
                   if (TopNode.SelectSingleNode("c:txPr", NameSpaceManager) == null)
                   {
                       CreateNode("c:txPr/a:bodyPr");
                       CreateNode("c:txPr/a:lstStyle");
                   }
                   _font = new ExcelTextFont(NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", new string[] { "spPr", "txPr", "dLblPos", "showVal", "showCatName ", "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" });
               }
               return _font;
           }
       }
    }
}
