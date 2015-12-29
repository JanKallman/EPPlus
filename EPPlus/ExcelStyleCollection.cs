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
 * Jan Källman		    Initial Release		        2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using System.Linq;
namespace OfficeOpenXml
{
    /// <summary>
    /// Base collection class for styles.
    /// </summary>
    /// <typeparam name="T">The style type</typeparam>
    public class ExcelStyleCollection<T> : IEnumerable<T>
    {
        public ExcelStyleCollection()
        {
            _setNextIdManual = false;
        }
        bool _setNextIdManual;
        public ExcelStyleCollection(bool SetNextIdManual)
        {
            _setNextIdManual = SetNextIdManual;
        }
        public XmlNode TopNode { get; set; }
        internal List<T> _list = new List<T>();
        Dictionary<string, int> _dic = new Dictionary<string, int>(StringComparer.InvariantCultureIgnoreCase);
        internal int NextId=0;
        #region IEnumerable<T> Members

        public IEnumerator<T> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        #endregion
        public T this[int PositionID]
        {
            get
            {
                return _list[PositionID];
            }
        }
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
        //internal int Add(T item)
        //{
        //    _list.Add(item);
        //    if (_setNextIdManual) NextId++;
        //    return _list.Count-1;
        //}
        internal int Add(string key, T item)
        {
            _list.Add(item);
            if (!_dic.ContainsKey(key.ToLower(CultureInfo.InvariantCulture))) _dic.Add(key.ToLower(CultureInfo.InvariantCulture), _list.Count - 1);
            if (_setNextIdManual) NextId++;
            return _list.Count-1;
        }
        /// <summary>
        /// Finds the key 
        /// </summary>
        /// <param name="key">the key to be found</param>
        /// <param name="obj">The found object.</param>
        /// <returns>True if found</returns>
        internal bool FindByID(string key, ref T obj)
        {
            if (_dic.ContainsKey(key))
            {
                obj = _list[_dic[key.ToLower(CultureInfo.InvariantCulture)]];
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Find Index
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        internal int FindIndexByID(string key)
        {
            if (_dic.ContainsKey(key))
            {
                return _dic[key];
            }
            else
            {
                return int.MinValue;
            }
        }
        internal bool ExistsKey(string key)
        {
            return _dic.ContainsKey(key);
        }
        internal void Sort(Comparison<T> c)
        {
            _list.Sort(c);
        }
    }
}
