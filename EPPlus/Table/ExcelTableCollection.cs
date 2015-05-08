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
 * Jan Källman		Added		30-AUG-2010
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// A collection of table objects
    /// </summary>
    public class ExcelTableCollection : IEnumerable<ExcelTable>
    {
        List<ExcelTable> _tables = new List<ExcelTable>();
        internal Dictionary<string, int> _tableNames = new Dictionary<string, int>();
        ExcelWorksheet _ws;        
        internal ExcelTableCollection(ExcelWorksheet ws)
        {
            var pck = ws._package.Package;
            _ws = ws;
            foreach(XmlElement node in ws.WorksheetXml.SelectNodes("//d:tableParts/d:tablePart", ws.NameSpaceManager))
            {
                var rel = ws.Part.GetRelationship(node.GetAttribute("id",ExcelPackage.schemaRelationships));
                var tbl = new ExcelTable(rel, ws);
                _tableNames.Add(tbl.Name, _tables.Count);
                _tables.Add(tbl);
            }
        }
        private ExcelTable Add(ExcelTable tbl)
        {
            _tables.Add(tbl);
            _tableNames.Add(tbl.Name, _tables.Count - 1);
            if (tbl.Id >= _ws.Workbook._nextTableID)
            {
                _ws.Workbook._nextTableID = tbl.Id + 1;
            }
            return tbl;
        }

        /// <summary>
        /// Create a table on the supplied range
        /// </summary>
        /// <param name="Range">The range address including header and total row</param>
        /// <param name="Name">The name of the table. Must be unique </param>
        /// <returns>The table object</returns>
        public ExcelTable Add(ExcelAddressBase Range, string Name)
        {
            if (string.IsNullOrEmpty(Name))
            {
                Name = GetNewTableName();
            }
            else if (_ws.Workbook.ExistsTableName(Name))
            {
                throw (new ArgumentException("Tablename is not unique"));
            }
            foreach (var t in _tables)
            {
                if (t.Address.Collide(Range) != ExcelAddressBase.eAddressCollition.No)
                {
                    throw (new ArgumentException(string.Format("Table range collides with table {0}", t.Name)));
                }
            }
            return Add(new ExcelTable(_ws, Range, Name, _ws.Workbook._nextTableID));
        }

        internal string GetNewTableName()
        {
            string name = "Table1";
            int i = 2;
            while (_ws.Workbook.ExistsTableName(name))
            {
                name = string.Format("Table{0}", i++);
            }
            return name;
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _tables.Count;
            }
        }
        /// <summary>
        /// Get the table object from a range.
        /// </summary>
        /// <param name="Range">The range</param>
        /// <returns>The table. Null if no range matches</returns>
        public ExcelTable GetFromRange(ExcelRangeBase Range)
        {
            foreach (var tbl in Range.Worksheet.Tables)
            {
                if (tbl.Address._address == Range._address)
                {
                    return tbl;
                }
            }
            return null;
        }
        /// <summary>
        /// The table Index. Base 0.
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelTable this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _tables.Count)
                {
                    throw (new ArgumentOutOfRangeException("Table index out of range"));
                }
                return _tables[Index];
            }
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="Name">The name of the table</param>
        /// <returns>The table. Null if the table name is not found in the collection</returns>
        public ExcelTable this[string Name]
        {
            get
            {
                if (_tableNames.ContainsKey(Name))
                {
                    return _tables[_tableNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }
        public IEnumerator<ExcelTable> GetEnumerator()
        {
            return _tables.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _tables.GetEnumerator();
        }
    }
}
