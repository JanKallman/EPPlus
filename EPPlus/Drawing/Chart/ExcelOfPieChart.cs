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
    /// Provides access to ofpie-chart specific properties
    /// </summary>
    public class ExcelOfPieChart : ExcelPieChart
    {
        //internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node) :
        //    base(drawings, node)
        //{

        //}
        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot) :
            base(drawings, node, type, isPivot)
        {
                SetTypeProperties();
        }
        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
            SetTypeProperties();
        }

        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Zip.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
           base(drawings, node, uriChart, part, chartXml, chartNode)
        {
            SetTypeProperties();
        }

        private void SetTypeProperties()
        {
            if (ChartType == eChartType.BarOfPie)
            {
                OfPieType = ePieType.Bar;
            }
            else
            {
                OfPieType = ePieType.Pie;
            }
        }

        const string pieTypePath = "c:chartSpace/c:chart/c:plotArea/c:ofPieChart/c:ofPieType/@val";
        /// <summary>
        /// Type, pie or bar
        /// </summary>
        public ePieType OfPieType
        {
            get
            {
                if (_chartXmlHelper.GetXmlNodeString(pieTypePath) == "bar")
                    return ePieType.Bar;
                else
                {
                    return ePieType.Pie;
                }
            }
            internal set
            {
                _chartXmlHelper.CreateNode(pieTypePath,true);
                _chartXmlHelper.SetXmlNodeString(pieTypePath, value == ePieType.Bar ? "bar" : "pie");
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (name == "ofPieChart")
            {
                if (OfPieType==ePieType.Bar)
                {
                    return eChartType.BarOfPie;
                }
                else
                {
                    return eChartType.PieOfPie;
                }
            }
            return base.GetChartType(name);
        }
    }
}
