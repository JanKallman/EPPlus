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
 * Jan Källman		Added this class		        2010-01-28
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;

namespace OfficeOpenXml
{
    /// <summary>
    /// Collection for named ranges
    /// </summary>
    public class ExcelNamedRangeCollection : IEnumerable<ExcelNamedRange>
    {
        internal ExcelWorksheet _ws;
        internal ExcelWorkbook _wb;
        internal ExcelNamedRangeCollection(ExcelWorkbook wb)
        {
            _wb = wb;
            _ws = null;
        }
        internal ExcelNamedRangeCollection(ExcelWorkbook wb, ExcelWorksheet ws)
        {
            _wb = wb;
            _ws = ws;
        }
        List<ExcelNamedRange> _list = new List<ExcelNamedRange>();
        Dictionary<string, int> _dic = new Dictionary<string, int>();
        /// <summary>
        /// Add a new named range
        /// </summary>
        /// <param name="Name">The name</param>
        /// <param name="Range">The range</param>
        /// <returns></returns>
        public ExcelNamedRange Add(string Name, ExcelRangeBase Range)
        {
            ExcelNamedRange item;
            if (Range.IsName)
            {

                item = new ExcelNamedRange(Name, _wb,_ws, _dic.Count);
            }
            else
            {
                item = new ExcelNamedRange(Name, _ws, Range.Worksheet, Range.Address, _dic.Count);
            }

            AddName(Name, item);

            return item;
        }

        private void AddName(string Name, ExcelNamedRange item)
        {
            _dic.Add(Name.ToLower(), _list.Count);
            _list.Add(item);
        }
        /// <summary>
        /// Add a defined name referencing value
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public ExcelNamedRange AddValue(string Name, object value)
        {
            var item = new ExcelNamedRange(Name,_wb, _ws, _dic.Count);
            item.NameValue = value;
            AddName(Name, item);
            return item;
        }

        /// <summary>
        /// Add a defined name referencing a formula -- the method name contains a typo.
        /// This method is obsolete and will be removed in the future.
        /// Use <see cref="AddFormula"/>
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="Formula"></param>
        /// <returns></returns>
        [Obsolete("Call AddFormula() instead.  See Issue Tracker Id #14687")]
        public ExcelNamedRange AddFormla(string Name, string Formula)
        {
            return  this.AddFormula(Name, Formula);
        }

        /// <summary>
        /// Add a defined name referencing a formula
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="Formula"></param>
        /// <returns></returns>
        public ExcelNamedRange AddFormula(string Name, string Formula)
        {
            var item = new ExcelNamedRange(Name, _wb, _ws, _dic.Count);
            item.NameFormula = Formula;
            AddName(Name, item);
            return item;
        }
        /// <summary>
        /// Remove a defined name from the collection
        /// </summary>
        /// <param name="Name">The name</param>
        public void Remove(string Name)
        {
            Name=Name.ToLower();
            if(_dic.ContainsKey(Name))
            {
                var ix = _dic[Name];

                for (int i = ix+1; i < _list.Count; i++)
                {
                    _dic.Remove(_list[i].Name.ToLower());
                    _list[i].Index--;
                    _dic.Add(_list[i].Name.ToLower(), _list[i].Index);
                }
                _dic.Remove(Name);
                _list.RemoveAt(ix);
            }
        }
        /// <summary>
        /// Checks collection for the presence of a key
        /// </summary>
        /// <param name="key">key to search for</param>
        /// <returns>true if the key is in the collection</returns>
        public bool ContainsKey(string key)
        {
            return _dic.ContainsKey(key.ToLower());
        }
        /// <summary>
        /// The current number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _dic.Count;
            }
        }
        /// <summary>
        /// Name indexer
        /// </summary>
        /// <param name="Name">The name (key) for a Named range</param>
        /// <returns>a reference to the range</returns>
        /// <remarks>
        /// Throws a KeyNotFoundException if the key is not in the collection.
        /// </remarks>
        public ExcelNamedRange this[string Name]
        {
            get
            {
                return _list[_dic[Name.ToLower()]];
            }
        }
        public ExcelNamedRange this[int Index]
        {
            get
            {
                return _list[Index];
            }
        }

        #region "IEnumerable"
        #region IEnumerable<ExcelNamedRange> Members
        /// <summary>
        /// Implement interface method IEnumerator&lt;ExcelNamedRange&gt; GetEnumerator()
        /// </summary>
        /// <returns></returns>
        public IEnumerator<ExcelNamedRange> GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        #endregion
        #region IEnumerable Members
        /// <summary>
        /// Implement interface method IEnumeratable GetEnumerator()
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
        #endregion
    }
}
