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
using System.IO.Packaging;
using System.Collections;

namespace OfficeOpenXml.Drawing.Chart
{
   /// <summary>
   /// A chart serie
   /// </summary>
    public class ExcelChartSerie : XmlHelper
   {
       internal ExcelChartSeries _chartSeries;
       protected XmlNode _node;
       protected XmlNamespaceManager _ns;
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chartSeries">Parent collection</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
       /// <param name="isPivot">Is pivotchart</param>
       internal ExcelChartSerie(ExcelChartSeries chartSeries, XmlNamespaceManager ns, XmlNode node, bool isPivot)
           : base(ns,node)
       {
           _chartSeries = chartSeries;
           _node=node;
           _ns=ns;
           SchemaNodeOrder = new string[] { "idx", "order", "tx", "marker","trendline", "explosion", "dLbls", "cat", "val", "yVal","xVal", "smooth" };

           if (chartSeries.Chart.ChartType == eChartType.XYScatter ||
               chartSeries.Chart.ChartType == eChartType.XYScatterLines ||
               chartSeries.Chart.ChartType == eChartType.XYScatterLinesNoMarkers ||
               chartSeries.Chart.ChartType == eChartType.XYScatterSmooth ||
               chartSeries.Chart.ChartType == eChartType.XYScatterSmoothNoMarkers)
           {
               _seriesTopPath = "c:yVal";
               _xSeriesTopPath = "c:xVal";
           }
           else
           {
               _seriesTopPath = "c:val";
               _xSeriesTopPath = "c:cat";
           }
           _seriesPath = string.Format(_seriesPath, _seriesTopPath);
           _xSeriesPath = string.Format(_xSeriesPath, _xSeriesTopPath, isPivot ? "c:multiLvlStrRef" : "c:numRef");
       }
       internal void SetID(string id)
       {
           SetXmlNodeString("c:idx/@val",id);
           SetXmlNodeString("c:order/@val", id);
       }
       const string headerPath="c:tx/c:v";
       /// <summary>
       /// Header for the serie.
       /// </summary>
       public string Header 
       {
           get
           {
                return GetXmlNodeString(headerPath);
            }
            set
            {
                Cleartx();
                SetXmlNodeString(headerPath, value);            
            }
        }

       private void Cleartx()
       {
           var n = TopNode.SelectSingleNode("c:tx", NameSpaceManager);
           if (n != null)
           {
               n.InnerXml = "";
           }
       }
       const string headerAddressPath = "c:tx/c:strRef/c:f";
        /// <summary>
       /// Header address for the serie.
       /// </summary>
       public ExcelAddressBase HeaderAddress
       {
           get
           {
               string address = GetXmlNodeString(headerAddressPath);
               if (address == "")
               {
                   return null;
               }
               else
               {
                   return new ExcelAddressBase(address);
               }
            }
            set
            {
                if (value._fromCol != value._toCol || value._fromRow != value._toRow || value.Addresses != null)
                {
                    throw (new Exception("Address must be a single cell"));
                }

                Cleartx();
                SetXmlNodeString(headerAddressPath, ExcelCell.GetFullAddress(value.WorkSheet, value.Address));
                SetXmlNodeString("c:tx/c:strRef/c:strCache/c:ptCount/@val", "0");
            }
        }        
        string _seriesTopPath;
        string _seriesPath = "{0}/c:numRef/c:f";       
       /// <summary>
       /// Set this to a valid address or the drawing will be invalid.
       /// </summary>
       public string Series
       {
           get
           {
               return GetXmlNodeString(_seriesPath);
           }
           set
           {
               if (_chartSeries.Chart.ChartType == eChartType.Bubble)
               {
                   throw(new Exception("Bubble charts is not supported yet"));
               }
               CreateNode(_seriesPath,true);
               SetXmlNodeString(_seriesPath, ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));
               
               XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache",_seriesTopPath), _ns);
               if (cache != null)
               {
                   cache.ParentNode.RemoveChild(cache);
               }

               if (_chartSeries.Chart.PivotTableSource != null)
               {
                   SetXmlNodeString(string.Format("{0}/c:numRef/c:numCache", _seriesTopPath), "General");
               }
               
               XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:numLit",_seriesTopPath), _ns);
               if (lit != null)
               {
                   lit.ParentNode.RemoveChild(lit);
               }
           }

       }
       string _xSeriesTopPath;
       string _xSeriesPath = "{0}/{1}/c:f";
       /// <summary>
       /// Set an address for the horisontal labels
       /// </summary>
       public string XSeries
       {
           get
           {
               return GetXmlNodeString(_xSeriesPath);
           }
           set
           {
               CreateNode(_xSeriesPath, true);
               SetXmlNodeString(_xSeriesPath, ExcelCellBase.GetFullAddress(_chartSeries.Chart.WorkSheet.Name, value));

               XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numRef/c:numCache",_xSeriesTopPath), _ns);
               if (cache != null)
               {
                   cache.ParentNode.RemoveChild(cache);
               }

               XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:numLit",_xSeriesTopPath), _ns);
               if (lit != null)
               {
                   lit.ParentNode.RemoveChild(lit);
               }
           }
       }
       ExcelChartTrendlineCollection _trendLines = null;
       /// <summary>
       /// Access to the trendline collection
       /// </summary>
        public ExcelChartTrendlineCollection TrendLines
        {
            get
            {
                if (_trendLines == null)
                {
                    _trendLines = new ExcelChartTrendlineCollection(this);
                }
                return _trendLines;
            }
        }
   }
}
