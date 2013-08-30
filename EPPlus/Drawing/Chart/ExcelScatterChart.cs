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
 * Jan Källman		Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to scatter chart specific properties
    /// </summary>
    public sealed class ExcelScatterChart : ExcelChart
    {
        internal ExcelScatterChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
            SetTypeProperties();
        }

        internal ExcelScatterChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
            base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            SetTypeProperties();
        }

        internal ExcelScatterChart(ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
            SetTypeProperties();
        }
        private void SetTypeProperties()
        {
           /***** ScatterStyle *****/
           if(ChartType == eChartType.XYScatter ||
              ChartType == eChartType.XYScatterLines ||
              ChartType == eChartType.XYScatterLinesNoMarkers)
           {
               ScatterStyle = eScatterStyle.LineMarker;
          }
           else if (
              ChartType == eChartType.XYScatterSmooth ||
              ChartType == eChartType.XYScatterSmoothNoMarkers) 
           {
               ScatterStyle = eScatterStyle.SmoothMarker;
           }
        }
        #region "Grouping Enum Translation"
        string _scatterTypePath = "c:scatterStyle/@val";
        private eScatterStyle GetScatterEnum(string text)
        {
            switch (text)
            {
                case "smoothMarker":
                    return eScatterStyle.SmoothMarker;
                default:
                    return eScatterStyle.LineMarker;
            }
        }

        private string GetScatterText(eScatterStyle shatterStyle)
        {
            switch (shatterStyle)
            {
                case eScatterStyle.SmoothMarker:
                    return "smoothMarker";
                default:
                    return "lineMarker";
            }
        }
        #endregion
        /// <summary>
        /// If the scatter has LineMarkers or SmoothMarkers
        /// </summary>
        public eScatterStyle ScatterStyle
        {
            get
            {
                return GetScatterEnum(_chartXmlHelper.GetXmlNodeString(_scatterTypePath));
            }
            internal set
            {
                _chartXmlHelper.CreateNode(_scatterTypePath, true);
                _chartXmlHelper.SetXmlNodeString(_scatterTypePath, GetScatterText(value));
            }
        }
        string MARKER_PATH = "c:marker/@val";
        /// <summary>
        /// If the series has markers
        /// </summary>
        public bool Marker
        {
            get
            {
                return GetXmlNodeBool(MARKER_PATH, false);
            }
            set
            {
                SetXmlNodeBool(MARKER_PATH, value, false);
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (name == "scatterChart")
            {
                if (ScatterStyle==eScatterStyle.LineMarker)
                {
                    if (((ExcelScatterChartSerie)Series[0]).Marker == eMarkerStyle.None)
                    {
                        return eChartType.XYScatterLinesNoMarkers;
                    }
                    else
                    {
                        if(ExistNode("c:ser/c:spPr/a:ln/noFill"))
                        {
                            return eChartType.XYScatter;
                        }
                        else
                        {
                            return eChartType.XYScatterLines;
                        }
                    }
                }
                else if (ScatterStyle == eScatterStyle.SmoothMarker)
                {
                    if (((ExcelScatterChartSerie)Series[0]).Marker == eMarkerStyle.None)
                    {
                        return eChartType.XYScatterSmoothNoMarkers;
                    }
                    else
                    {
                        return eChartType.XYScatterSmooth;
                    }
                }
            }
            return base.GetChartType(name);
        }

    }
}
