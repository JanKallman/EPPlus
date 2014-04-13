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
    /// Provides access to pie chart specific properties
    /// </summary>
    public class ExcelPieChart : ExcelChart
    {
        internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot) :
            base(drawings, node, type, isPivot)
        {
        }
        internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
            base(drawings, node, type, topChart, PivotTableSource)
        {
        }

        internal ExcelPieChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
           base(drawings, node, uriChart, part, chartXml, chartNode)
        {
        }

        internal ExcelPieChart(ExcelChart topChart, XmlNode chartNode) :
            base(topChart, chartNode)
        {
        }
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
            if (name == "pieChart")
            {
                if (Series.Count > 0 && ((ExcelPieChartSerie)Series[0]).Explosion>0)
                {
                    return eChartType.PieExploded;
                }
                else
                {
                    return eChartType.Pie;
                }
            }
            else if (name == "pie3DChart")
            {
                if (Series.Count > 0 && ((ExcelPieChartSerie)Series[0]).Explosion > 0)
                {
                    return eChartType.PieExploded3D;
                }
                else
                {
                    return eChartType.Pie3D;
                }
            }
            return base.GetChartType(name);
        }
    }
}
