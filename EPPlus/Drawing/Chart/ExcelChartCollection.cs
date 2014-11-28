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

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Enumerates charttypes 
    /// </summary>
    public class ExcelChartCollection : IEnumerable<ExcelChart>
    {
        List<ExcelChart> _list = new List<ExcelChart>();
        ExcelChart _topChart;
        internal ExcelChartCollection(ExcelChart chart)
        {
            _topChart = chart;
            _list.Add(chart);
        }
        internal void Add(ExcelChart chart)
        {
            _list.Add(chart);
        }
        /// <summary>
        /// Add a new charttype to the chart
        /// </summary>
        /// <param name="chartType">The type of the new chart</param>
        /// <returns></returns>
        public ExcelChart Add(eChartType chartType)
        {
            if (_topChart.PivotTableSource != null)
            {
                throw (new InvalidOperationException("Can not add other charttypes to a pivot chart"));
            }
            else if (ExcelChart.IsType3D(chartType) || _list[0].IsType3D())
            {
                throw(new InvalidOperationException("3D charts can not be combined with other charttypes"));
            }

            var prependingChartNode = _list[_list.Count - 1].TopNode;
            ExcelChart chart = ExcelChart.GetNewChart(_topChart.WorkSheet.Drawings, _topChart.TopNode, chartType, _topChart, null);

            _list.Add(chart);
            return chart;
        }
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
        IEnumerator<ExcelChart> IEnumerable<ExcelChart>.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        /// <summary>
        /// Returns a chart at the specific position.  
        /// </summary>
        /// <param name="PositionID">The position of the chart. 0-base</param>
        /// <returns></returns>
        public ExcelChart this[int PositionID]
        {
            get
            {
                return (_list[PositionID]);
            }
        }


}
}
