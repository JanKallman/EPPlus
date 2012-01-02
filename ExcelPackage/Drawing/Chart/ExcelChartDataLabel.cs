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
    /// Datalabel on chart level. 
    /// This class is inherited by ExcelChartSerieDataLabel
    /// </summary>
    public class ExcelChartDataLabel : XmlHelper
    {
       internal ExcelChartDataLabel(XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           XmlNode topNode = node.SelectSingleNode("c:dLbls", NameSpaceManager);
           if (topNode == null)
           {
               topNode = node.OwnerDocument.CreateElement("c", "dLbls", ExcelPackage.schemaChart);
               //node.InsertAfter(_topNode, node.SelectSingleNode("c:order", NameSpaceManager));
               InserAfter(node, "c:marker,c:tx,c:order,c:ser", topNode);
               SchemaNodeOrder = new string[] { "spPr", "txPr", "dLblPos", "showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize", "separator", "showLeaderLines" };
               topNode.InnerXml = "<c:showLegendKey val=\"0\" /><c:showVal val=\"0\" /><c:showCatName val=\"0\" /><c:showSerName val=\"0\" /><c:showPercent val=\"0\" /><c:showBubbleSize val=\"0\" /> <c:separator>\r\n</c:separator><c:showLeaderLines val=\"0\" />";                      
           }
           TopNode = topNode;
       }
       #region "Public properties"
       const string showValPath = "c:showVal/@val";
       /// <summary>
       /// Show the values 
       /// </summary>
        public bool ShowValue
       {
           get
           {
               return GetXmlNodeBool(showValPath);
           }
           set
           {
               SetXmlNodeString(showValPath, value ? "1" : "0");
           }
       }
       const string showCatPath = "c:showCatName/@val";
       /// <summary>
       /// Show category names
       /// </summary>
        public bool ShowCategory
       {
           get
           {
               return GetXmlNodeBool(showCatPath);
           }
           set
           {
               SetXmlNodeString(showCatPath, value ? "1" : "0");
           }
       }
       const string showSerPath = "c:showSerName/@val";
       /// <summary>
       /// Show series names
       /// </summary>
        public bool ShowSeriesName
       {
           get
           {
               return GetXmlNodeBool(showSerPath);
           }
           set
           {
               SetXmlNodeString(showSerPath, value ? "1" : "0");
           }
       }
       const string showPerentPath = "c:showPercent/@val";
       /// <summary>
       /// Show percent values
       /// </summary>
        public bool ShowPercent
       {
           get
           {
               return GetXmlNodeBool(showPerentPath);
           }
           set
           {
               SetXmlNodeString(showPerentPath, value ? "1" : "0");
           }
       }
       const string showLeaderLinesPath = "c:showLeaderLines/@val";
       /// <summary>
       /// Show the leader lines
       /// </summary>
        public bool ShowLeaderLines
       {
           get
           {
               return GetXmlNodeBool(showLeaderLinesPath);
           }
           set
           {
               SetXmlNodeString(showLeaderLinesPath, value ? "1" : "0");
           }
       }
       const string showBubbleSizePath = "c:showBubbleSize/@val";
       /// <summary>
       /// Bubble Size.
       /// </summary>
        public bool ShowBubbleSize
       {
           get
           {
               return GetXmlNodeBool(showBubbleSizePath);
           }
           set
           {
               SetXmlNodeString(showBubbleSizePath, value ? "1" : "0");
           }
       }
       const string showLegendKeyPath = "c:showLegendKey/@val";
       public bool ShowLegendKey
       {
           get
           {
               return GetXmlNodeBool(showLegendKeyPath);
           }
           set
           {
               SetXmlNodeString(showLegendKeyPath, value ? "1" : "0");
           }
       }
       const string separatorPath = "c:separator";
       /// <summary>
       /// Separator string 
       /// </summary>
        public string Separator
       {
           get
           {
               return GetXmlNodeString(separatorPath);
           }
           set
           {
               if (string.IsNullOrEmpty(value))
               {
                   DeleteNode(separatorPath);
               }
               else
               {
                   SetXmlNodeString(separatorPath, value);
               }
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
       string[] _paragraphSchemaOrder = new string[] { "spPr", "txPr", "dLblPos", "showVal", "showCatName", "showSerName", "showPercent", "separator", "showLeaderLines", "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" };
       ExcelTextFont _font = null;
       /// <summary>
       /// Access font properties
       /// </summary>
       public ExcelTextFont Font
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
                   _font = new ExcelTextFont(NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", _paragraphSchemaOrder);
               }
               return _font;
           }
       }
       #endregion
       #region "Position Enum Translation"
       protected string GetPosText(eLabelPosition pos)
       {
           switch (pos)
           {
               case eLabelPosition.Bottom:
                   return "b";
               case eLabelPosition.Center:
                   return "ctr";
               case eLabelPosition.InBase:
                   return "inBase";
               case eLabelPosition.InEnd:
                   return "inEnd";
               case eLabelPosition.Left:
                   return "l";
               case eLabelPosition.Right:
                   return "r";
               case eLabelPosition.Top:
                   return "t";
               case eLabelPosition.OutEnd:
                   return "outEnd";
               default:
                   return "bestFit";
           }
       }

       protected eLabelPosition GetPosEnum(string pos)
       {
           switch (pos)
           {
               case "b":
                   return eLabelPosition.Bottom;
               case "ctr":
                   return eLabelPosition.Center;
               case "inBase":
                   return eLabelPosition.InBase;
               case "inEnd":
                   return eLabelPosition.InEnd;
               case "l":
                   return eLabelPosition.Left;
               case "r":
                   return eLabelPosition.Right;
               case "t":
                   return eLabelPosition.Top;
               case "outEnd":
                   return eLabelPosition.OutEnd;
               default:
                   return eLabelPosition.BestFit;
           }
       }
    #endregion
    }
}
