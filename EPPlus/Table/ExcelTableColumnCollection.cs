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
using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// A collection of table columns
    /// </summary>
    public class ExcelTableColumnCollection : IEnumerable<ExcelTableColumn>
    {
        List<ExcelTableColumn> _cols = new List<ExcelTableColumn>();
        Dictionary<string, int> _colNames = new Dictionary<string, int>(StringComparer.InvariantCultureIgnoreCase);
        public ExcelTableColumnCollection(ExcelTable table)
        {
            Table = table;
            foreach(XmlNode node in table.TableXml.SelectNodes("//d:table/d:tableColumns/d:tableColumn",table.NameSpaceManager))
            {                
                _cols.Add(new ExcelTableColumn(table.NameSpaceManager, node, table, _cols.Count));
                _colNames.Add(_cols[_cols.Count - 1].Name, _cols.Count - 1);
            }
        }
        /// <summary>
        /// A reference to the table object
        /// </summary>
        public ExcelTable Table
        {
            get;
            private set;
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _cols.Count;
            }
        }
        /// <summary>
        /// The column Index. Base 0.
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelTableColumn this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _cols.Count)
                {
                    throw (new ArgumentOutOfRangeException("Column index out of range"));
                }
                return _cols[Index] as ExcelTableColumn;
            }
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="Name">The name of the table</param>
        /// <returns>The table column. Null if the table name is not found in the collection</returns>
        public ExcelTableColumn this[string Name]
        {
            get
            {
                if (_colNames.ContainsKey(Name))
                {
                    return _cols[_colNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }

        IEnumerator<ExcelTableColumn> IEnumerable<ExcelTableColumn>.GetEnumerator()
        {
            return _cols.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _cols.GetEnumerator();
        }
        internal string GetUniqueName(string name)
        {            
            if (_colNames.ContainsKey(name))
            {
                var newName = name;
                var i = 2;
                do
                {
                    newName = name+(i++).ToString(CultureInfo.InvariantCulture);
                }
                while (_colNames.ContainsKey(newName));
                return newName;
            }
            return name;
        }
    }
}
