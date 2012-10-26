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
using System.Globalization;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to doughnut chart specific properties
    /// </summary>
    public class ExcelDoughnutChart : ExcelPieChart
    {
        //internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node) :
        //    base(drawings, node)
        //{
        //    SetPaths();
        //}
        internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot) :
            base(drawings, node, type, isPivot)
        {
            //SetPaths();
        }
        internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
            //SetPaths();
        }
        internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Zip.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
           base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            //SetPaths();
        }

        internal ExcelDoughnutChart(ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
            //SetPaths();
        }

        //private void SetPaths()
        //{
        //    string chartNodeText = GetChartNodeText();
        //    _firstSliceAngPath = string.Format(_firstSliceAngPath, chartNodeText);
        //    _holeSizePath = string.Format(_holeSizePath, chartNodeText);
        //}
        //string _firstSliceAngPath = "c:chartSpace/c:chart/c:plotArea/{0}/c:firstSliceAng/@val";
        string _firstSliceAngPath = "c:firstSliceAng/@val";
        /// <summary>
        /// Angle of the first slize
        /// </summary>
        public decimal FirstSliceAngle
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeDecimal(_firstSliceAngPath);
            }
            internal set
            {
                _chartXmlHelper.SetXmlNodeString(_firstSliceAngPath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        //string _holeSizePath = "c:chartSpace/c:chart/c:plotArea/{0}/c:holeSize/@val";
        string _holeSizePath = "c:holeSize/@val";
        /// <summary>
        /// Size of the doubnut hole
        /// </summary>
        public decimal HoleSize
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeDecimal(_holeSizePath);
            }
            internal set
            {
                _chartXmlHelper.SetXmlNodeString(_holeSizePath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (name == "doughnutChart")
            {
                if (((ExcelPieChartSerie)Series[0]).Explosion > 0)
                {
                    return eChartType.DoughnutExploded;
                }
                else
                {
                    return eChartType.Doughnut;
                }
            }
            return base.GetChartType(name);
        }
    }
}
