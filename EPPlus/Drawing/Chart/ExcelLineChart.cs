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
    /// Provides access to line chart specific properties
    /// </summary>
    public class ExcelLineChart : ExcelChart
    {
        #region "Constructors"
        internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
            base(drawings, node, uriChart, part, chartXml, chartNode)
        {
        }

        internal ExcelLineChart (ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
        }
        internal ExcelLineChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
            Smooth = false;
        }
        #endregion
        string MARKER_PATH="c:marker/@val";
        /// <summary>
        /// If the series has markers
        /// </summary>
        public bool Marker
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(MARKER_PATH, false);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(MARKER_PATH, value, false);
            }
        }

        string SMOOTH_PATH = "c:smooth/@val";
        /// <summary>
        /// If the series has smooth lines
        /// </summary>
        public bool Smooth
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(SMOOTH_PATH, false);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(SMOOTH_PATH, value);
            }
        }
        //string _chartTopPath = "c:chartSpace/c:chart/c:plotArea/{0}";
        ExcelChartDataLabel _DataLabel = null;
        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartDataLabel(NameSpaceManager, ChartNode);
                }
                return _DataLabel;
            }
        }
        internal override eChartType GetChartType(string name)
        {
               if(name=="lineChart")
               {
                   if(Marker)
                   {
                       if(Grouping==eGrouping.Stacked)
                       {
                           return eChartType.LineMarkersStacked;
                       }
                       else if (Grouping == eGrouping.PercentStacked)
                       {
                           return eChartType.LineMarkersStacked100;
                       }
                       else
                       {
                           return eChartType.LineMarkers;
                       }
                   }
                   else
                   {
                       if(Grouping==eGrouping.Stacked)
                       {
                           return eChartType.LineStacked;
                       }
                       else if (Grouping == eGrouping.PercentStacked)
                       {
                           return eChartType.LineStacked100;
                       }
                       else
                       {
                           return eChartType.Line;
                       }
                   }
               }
               else if (name=="line3DChart")
               {
                   return eChartType.Line3D;               
               }
               return base.GetChartType(name);
        }
    }
}
