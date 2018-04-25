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
 * Jan Källman		Added		2017-09-20
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Sparkline
{
    /// <summary>
    /// A collection of sparkline groups
    /// </summary>
    public class ExcelSparklineGroupCollection : IEnumerable<ExcelSparklineGroup>
    {
        ExcelWorksheet _ws;
        List<ExcelSparklineGroup> _lst;
        internal ExcelSparklineGroupCollection(ExcelWorksheet ws)
        {
            _ws = ws;
            _lst = new List<ExcelSparklineGroup>();
            LoadSparklines();
        }
        const string _extPath = "/d:worksheet/d:extLst/d:ext";
        const string _searchPath = "[@uri='{05C60535-1F16-4fd2-B633-F4F36F0B64E0}']";
        const string _topSearchPath = _extPath + _searchPath + "/x14:sparklineGroups";
        const string _topPath = _extPath + "/x14:sparklineGroups";

        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _lst.Count;
            }            
        }
        /// <summary>
        /// Adds a new sparklinegroup to the collection
        /// </summary>
        /// <param name="type">Type of sparkline</param>
        /// <param name="locationRange">The location of the sparkline group. The range must have one row or column and must match the number of rows/columns in the datarange</param>
        /// <param name="dataRange">The data for the sparkline group</param>
        /// <returns></returns>
        public ExcelSparklineGroup Add(eSparklineType type, ExcelAddressBase locationRange, ExcelAddressBase dataRange)
        {
            if(locationRange.Rows==1)
            {
                if(locationRange.Columns==dataRange.Rows)
                {
                    return AddGroup(type, locationRange, dataRange, true);
                }
                else if(locationRange.Columns== dataRange.Columns)
                {
                    return AddGroup(type, locationRange, dataRange, false);
                }
                else
                {
                    throw (new ArgumentException("dataRange is not valid. dataRange columns or rows must match number of rows in locationRange"));
                }
            }
            else if(locationRange.Columns==1)
            {
                if (locationRange.Rows== dataRange.Columns)
                {
                    return AddGroup(type, locationRange, dataRange, false);
                }
                else if (locationRange.Rows== dataRange.Rows)
                {
                    return AddGroup(type, locationRange, dataRange, true);
                }
                else
                {
                    throw (new ArgumentException("dataRange is not valid. dataRange columns or rows must match number of columns in locationRange"));
                }
            }
            else
            {
                throw (new ArgumentException("locationRange is not valid. Range must be one Column or Row only"));
            }
        }

        private ExcelSparklineGroup AddGroup(eSparklineType type, ExcelAddressBase locationRange, ExcelAddressBase dataRange, bool isRows)
        {
            var group = NewSparklineGroup();
            group.Type = type;
            var row = locationRange._fromRow;
            var col = locationRange._fromCol;

            var drFromRow = dataRange._fromRow;
            var drFromCol = dataRange._fromCol;
            var drToRow = isRows ? dataRange._fromRow : dataRange._toRow;
            var drToCol = isRows ? dataRange._toCol : dataRange._fromCol;

            var cells = (locationRange._fromRow==locationRange._toRow ? locationRange._toCol - locationRange._fromCol: locationRange._toRow- locationRange._fromRow)+1;
            var cell = 0;

            while (cell<cells)
            {
                var f = new ExcelCellAddress(row, col);
                var sqref = new ExcelAddressBase(dataRange.WorkSheet, drFromRow, drFromCol, drToRow, drToCol);
                group.Sparklines.Add(f, dataRange.WorkSheet, sqref);
                cell++;
                if(locationRange._fromRow == locationRange._toRow)
                {
                    col++;
                }
                else
                {
                    row++;
                }
                if(isRows)
                {
                    drFromRow++;
                    drToRow++;
                }
                else
                {
                    drFromCol++;
                    drToCol++;
                }
            }

            group.ColorSeries.Rgb = "FF376092";
            group.ColorNegative.Rgb = "FFD00000";
            group.ColorMarkers.Rgb = "FFD00000";
            group.ColorAxis.Rgb = "FF000000";
            group.ColorFirst.Rgb = "FFD00000";
            group.ColorLast.Rgb = "FFD00000";
            group.ColorHigh.Rgb = "FFD00000";
            group.ColorLow.Rgb = "FFD00000";
            return group;
        }

        private ExcelSparklineGroup NewSparklineGroup()
        {
            var xh = new XmlHelperInstance(_ws.NameSpaceManager, _ws.WorksheetXml); //SelectSingleNode("/d:worksheet", _ws.NameSpaceManager)
            if (!xh.ExistNode(_extPath + _searchPath))
            {
                var ext = xh.CreateNode(_extPath, false, true);
                if (ext.Attributes["uri"] == null)
                {
                    ((XmlElement)ext).SetAttribute("uri", "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}");        //Guid URI for sparklines... https://msdn.microsoft.com/en-us/library/dd905242(v=office.12).aspx
                }
            }
            var parent = xh.CreateNode(_topSearchPath);

            var topNode = _ws.WorksheetXml.CreateElement("x14","sparklineGroup", ExcelPackage.schemaMainX14);
            topNode.SetAttribute("xmlns:xm", ExcelPackage.schemaMainXm);
            topNode.SetAttribute("uid", ExcelPackage.schemaXr2, $"{{{Guid.NewGuid().ToString()}}}");
            parent.AppendChild(topNode);
            topNode.AppendChild(topNode.OwnerDocument.CreateElement("x14", "sparklines", ExcelPackage.schemaMainX14));
            return new ExcelSparklineGroup(_ws.NameSpaceManager, topNode, _ws);
        }

        private void LoadSparklines()
        {
            var grps=_ws.WorksheetXml.SelectNodes(_topPath + "/x14:sparklineGroup", _ws.NameSpaceManager);
            foreach (XmlElement grp in grps)
            {
                _lst.Add(new ExcelSparklineGroup(_ws.NameSpaceManager, grp, _ws));
            }
        }
        /// <summary>
        /// Returns the sparklinegroup at the specified position.  
        /// </summary>
        /// <param name="index">The position of the Sparklinegroup. 0-base</param>
        /// <returns></returns>
        public ExcelSparklineGroup this[int index]
        {
            get
            {
                return (_lst[index]);
            }
        }

        public IEnumerator<ExcelSparklineGroup> GetEnumerator()
        {
            return _lst.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _lst.GetEnumerator();
        }
        /// <summary>
        /// Removes the sparkline.
        /// </summary>
        /// <param name="index">The index of the item to be removed</param>
        public void RemoveAt(int index)
        {
            Remove(_lst[index]);
        }
        /// <summary>
        /// Removes the sparkline.
        /// </summary>
        /// <param name="sparklineGroup">The sparklinegroup to be removed</param>
        public void Remove(ExcelSparklineGroup sparklineGroup)
        {
            sparklineGroup.TopNode.ParentNode.RemoveChild(sparklineGroup.TopNode);
            _lst.Remove(sparklineGroup);
        }
    }
}
