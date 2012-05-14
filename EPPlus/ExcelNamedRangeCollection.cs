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
        Dictionary<string, ExcelNamedRange> _dic = new Dictionary<string, ExcelNamedRange>();
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

                item = new ExcelNamedRange(Name, _wb,_ws);
            }
            else
            {
                item = new ExcelNamedRange(Name, _ws, Range.Worksheet, Range.Address);
            }
            
            _dic.Add(Name, item);
            return item;
        }
        /// <summary>
        /// Add a defined name referencing value
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public ExcelNamedRange AddValue(string Name, object value)
        {
            var item = new ExcelNamedRange(Name,_wb, _ws);
            item.NameValue = value;
            _dic.Add(Name, item);
            return item;
        }
        /// <summary>
        /// Add a defined name referencing a formula
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="Formula"></param>
        /// <returns></returns>
        public ExcelNamedRange AddFormla(string Name, string Formula)
        {
            var item = new ExcelNamedRange(Name, _wb, _ws);
            item.NameFormula = Formula;
            _dic.Add(Name, item);
            return item;
        }
        /// <summary>
        /// Remove a defined name from the collection
        /// </summary>
        /// <param name="Name">The name</param>
        public void Remove(string Name)
        {
            _dic.Remove(Name);
        }
        public bool ContainsKey(string key)
        {
            return _dic.ContainsKey(key);
        }
        public int Count
        {
            get
            {
                return _dic.Count;
            }
        }
        /// <summary>
        /// Names
        /// </summary>
        /// <param name="Name">The name</param>
        /// <returns></returns>
        public ExcelNamedRange this[string Name]
        {
            get
            {
                return _dic[Name];
            }
        }

        #region "IEnumerable"
        #region IEnumerable<ExcelNamedRange> Members
        public IEnumerator<ExcelNamedRange> GetEnumerator()
        {
            return _dic.Values.GetEnumerator();
        }
        #endregion
        #region IEnumerable Members
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _dic.Values.GetEnumerator();
        }

        #endregion
        #endregion
    }
}
