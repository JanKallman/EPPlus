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
    /// Collection of sparklines
    /// </summary>
    public class ExcelSparklineCollection : IEnumerable<ExcelSparkline>
    {
        ExcelSparklineGroup _slg;
        List<ExcelSparkline> _lst;
        internal ExcelSparklineCollection(ExcelSparklineGroup slg)
        {
            _slg = slg;
            _lst = new List<ExcelSparkline>();
            LoadSparklines();
        }
        const string _topPath = "x14:sparklines/x14:sparkline";
        /// <summary>
        /// Number of sparklines in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _lst.Count;
            }            
        }

        private void LoadSparklines()
        {
            var grps=_slg.TopNode.SelectNodes(_topPath, _slg.NameSpaceManager);
            foreach(XmlElement grp in grps)
            {
                _lst.Add(new ExcelSparkline(_slg.NameSpaceManager, grp));
            }
        }
        /// <summary>
        /// Returns the sparklinegroup at the specified position.  
        /// </summary>
        /// <param name="index">The position of the Sparklinegroup. 0-base</param>
        /// <returns></returns>
        public ExcelSparkline this[int index]
        {
            get
            {
                return (_lst[index]);
            }
        }

        public IEnumerator<ExcelSparkline> GetEnumerator()
        {
            return _lst.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _lst.GetEnumerator();
        }

        internal void Add(ExcelCellAddress cell, string worksheetName, ExcelAddressBase sqref)
        {
            var sparkline = _slg.TopNode.OwnerDocument.CreateElement("x14","sparkline", ExcelPackage.schemaMainX14);            
            var sls = _slg.TopNode.SelectSingleNode("x14:sparklines", _slg.NameSpaceManager);

            sls.AppendChild(sparkline);
            _slg.TopNode.AppendChild(sls);
            var sl = new ExcelSparkline(_slg.NameSpaceManager, sparkline);
            sl.Cell = cell;
            sl.RangeAddress = sqref;
            _lst.Add(sl);
        }
    }
}
