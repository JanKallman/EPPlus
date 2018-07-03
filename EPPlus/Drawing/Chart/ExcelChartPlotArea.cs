/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A charts plot area
    /// </summary>
    public sealed class ExcelChartPlotArea :  XmlHelper
    {
        ExcelChart _firstChart;
        internal ExcelChartPlotArea(XmlNamespaceManager ns, XmlNode node, ExcelChart firstChart)
           : base(ns,node)
       {
           _firstChart = firstChart;
            if (TopNode.SelectSingleNode("c:dTable", NameSpaceManager) != null)
            {
                _dataTable = new ExcelChartDataTable(NameSpaceManager, TopNode);
            }
        }

        ExcelChartCollection _chartTypes;
        public ExcelChartCollection ChartTypes
        {
            get
            {
                if (_chartTypes == null)
                {
                    _chartTypes = new ExcelChartCollection(_firstChart); 
                }
                return _chartTypes;
            }
        }
        #region Data table
        /// <summary>
        /// Creates a data table in the plotarea
        /// The datatable can also be accessed via the DataTable propery
        /// <see cref="DataTable"/>
        /// </summary>
        public ExcelChartDataTable CreateDataTable()
        {
            if(_dataTable!=null)
            {
                throw (new InvalidOperationException("Data table already exists"));
            }

            _dataTable = new ExcelChartDataTable(NameSpaceManager, TopNode);
            return _dataTable;
        }
        /// <summary>
        /// Remove the data table if it's created in the plotarea
        /// </summary>
        public void RemoveDataTable()
        {
            DeleteAllNode("c:dTable");
            _dataTable = null;
        }
        ExcelChartDataTable _dataTable = null;
        /// <summary>
        /// The data table object.
        /// Use the CreateDataTable method to create a datatable if it does not exist.
        /// <see cref="CreateDataTable"/>
        /// <see cref="RemoveDataTable"/>
        /// </summary>
        public ExcelChartDataTable DataTable
        {
            get
            {
                return _dataTable;
            }
        }
        #endregion
        ExcelDrawingFill _fill = null;
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
    }
}
