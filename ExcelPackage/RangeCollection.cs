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
 * Jan Källman		    Added       		        2010-02-04
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using OfficeOpenXml.Drawing.Vml;
namespace OfficeOpenXml
{
    /// <summary>
    /// This is the store for all Rows, Columns and Cells.
    /// It is a Dictionary implementation that allows you to change the Key (the RowID, ColumnID or CellID )
    /// </summary>
    internal class RangeCollection : IEnumerator<IRangeID>, IEnumerable
    {
        private class IndexItem
        {
            internal IndexItem(ulong cellId)
            {
                RangeID = cellId;
            }
            internal IndexItem(ulong cellId, int listPointer)
	        {
                RangeID = cellId;
                ListPointer=listPointer;
	        }
            internal ulong RangeID;
            internal int ListPointer;
        }
        /// <summary>
        /// Compares an IndexItem
        /// </summary>
        internal class Compare : IComparer<IndexItem>
        {
            #region IComparer<IndexItem> Members
            int IComparer<IndexItem>.Compare(IndexItem x, IndexItem y)
            {
                return x.RangeID < y.RangeID ? -1 : x.RangeID > y.RangeID ? 1 : 0;
            }

            #endregion
        }
        IndexItem[] _cellIndex;
        List<IRangeID> _cells;
        Compare _comparer;
        /// <summary>
        /// Creates a new collection
        /// </summary>
        /// <param name="cells">The Cells. This list must be sorted</param>
        internal RangeCollection(List<IRangeID> cells)
        {   
            _cells = cells;
            _comparer = new Compare();
            InitSize(_cells);
            for (int i = 0; i < _cells.Count; i++)
            {
                _cellIndex[i] = new IndexItem(cells[i].RangeID, i);
            }
        }
        /// <summary>
        /// Return the item with the RangeID
        /// </summary>
        /// <param name="RangeID"></param>
        /// <returns></returns>
        internal IRangeID this[ulong RangeID]
        {
            get
            {
                return _cells[_cellIndex[IndexOf(RangeID)].ListPointer];
            }
        }
        /// <summary>
        /// Return specified index from the sorted list
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        internal IRangeID this[int Index]
        {
            get
            {
                return _cells[_cellIndex[Index].ListPointer];
            }
        }
        internal int Count
        {
            get
            {
                return _cells.Count;
            }
        }
        internal void Add(IRangeID cell)
        {
            var ix = IndexOf(cell.RangeID);
            if (ix >= 0)
            {
                throw (new Exception("Item already exist"));
            }
            Insert(~ix, cell);
        }
        internal void Delete(ulong key)
        {
            var ix = IndexOf(key);
            if (ix < 0)
            {
                throw (new Exception("Key does not exist"));
            }
            int listPointer = _cellIndex[ix].ListPointer;
            Array.Copy(_cellIndex, ix + 1, _cellIndex, ix, _cells.Count - ix - 1);
            _cells.RemoveAt(listPointer);

            //Item is removed subtract one from all items with greater ListPointer
            for (int i = 0; i < _cells.Count; i++)
            {
                if (_cellIndex[i].ListPointer >= listPointer)
                {
                    _cellIndex[i].ListPointer--;
                }

            }
        }
        internal int IndexOf(ulong key)
        {
            return Array.BinarySearch<IndexItem>(_cellIndex, 0, _cells.Count, new IndexItem(key), _comparer);
        }
        internal bool ContainsKey(ulong key)
        {
            return IndexOf(key) < 0 ? false : true;
        }
        int _size { get; set; }
        #region "RangeID manipulation methods"
        /// <summary>
        /// Insert a number of rows in the collecion but dont update the cell only the index
        /// </summary>
        /// <param name="rowID"></param>
        /// <param name="rows"></param>
        /// <returns>Index of first rangeItem</returns>
        internal int InsertRowsUpdateIndex(ulong rowID, int rows)
        {
            int index = IndexOf(rowID);
            if (index < 0) index = ~index; //No match found invert to get start cell
            ulong rowAdd = (((ulong)rows) << 29);
            for (int i = index; i < _cells.Count; i++)
            {
                _cellIndex[i].RangeID += rowAdd;
            }
            return index;
        }
        /// <summary>
        /// Insert a number of rows in the collecion
        /// </summary>
        /// <param name="rowID"></param>
        /// <param name="rows"></param>
        /// <returns>Index of first rangeItem</returns>
        internal int InsertRows(ulong rowID, int rows)
        {
            int index = IndexOf(rowID);
            if (index < 0) index = ~index; //No match found invert to get start cell
            ulong rowAdd=(((ulong)rows) << 29);
            for (int i = index; i < _cells.Count; i++)
            {
                _cellIndex[i].RangeID += rowAdd;
                _cells[_cellIndex[i].ListPointer].RangeID += rowAdd;
            }
            return index;
        }
        /// <summary>
        /// Delete rows from the collecion
        /// </summary>
        /// <param name="rowID"></param>
        /// <param name="rows"></param>
        /// <param name="updateCells">Update range id's on cells</param>
        internal int DeleteRows(ulong rowID, int rows, bool updateCells)
        {
            ulong rowAdd = (((ulong)rows) << 29);
            var index = IndexOf(rowID);
            if (index < 0) index = ~index; //No match found invert to get start cell

            if (index >= _cells.Count || _cellIndex[index] == null) return -1;   //No row above this row
            while (index < _cells.Count && _cellIndex[index].RangeID < rowID + rowAdd)
            {
                Delete(_cellIndex[index].RangeID);
            }

            int updIndex = IndexOf(rowID + rowAdd);
            if (updIndex < 0) updIndex = ~updIndex; //No match found invert to get start cell

            for (int i = updIndex; i < _cells.Count; i++)
            {
                _cellIndex[i].RangeID -= rowAdd;                        //Change the index
                if (updateCells) _cells[_cellIndex[i].ListPointer].RangeID -= rowAdd;    //Change the cell/row or column object
            }
            return index;
        }
        internal void InsertColumn(ulong ColumnID, int columns)
        {
            throw (new Exception("Working on it..."));
        }
        internal void DeleteColumn(ulong ColumnID,int columns)
        {
            throw (new Exception("Working on it..."));
        }
        #endregion
        #region "Private Methods"
        /// <summary>
        /// Init the size starting from 128 items. Double the size until the list fits.
        /// </summary>
        /// <param name="_cells"></param>
        private void InitSize(List<IRangeID> _cells)
        {
            _size = 128;
            while (_cells.Count > _size) _size <<= 1;
            _cellIndex = new IndexItem[_size];
        }
        /// <summary>
        /// Check the size and double the size if out of bound
        /// </summary>
        private void CheckSize()
        {
            if (_cells.Count >= _size)
            {
                _size <<= 1;
                Array.Resize(ref _cellIndex, _size);
            }
        }
        private void Insert(int ix, IRangeID cell)
        {
            CheckSize();
            Array.Copy(_cellIndex, ix, _cellIndex, ix + 1, _cells.Count - ix);
            _cellIndex[ix] = new IndexItem(cell.RangeID, _cells.Count);
            _cells.Add(cell);
        }
        #endregion

        #region IEnumerator<IRangeID> Members

        IRangeID IEnumerator<IRangeID>.Current
        {
            get { throw new NotImplementedException(); }
        }

        #endregion

        #region IDisposable for the enumerator Members

        void IDisposable.Dispose()
        {
            _ix = -1;
        }

        #endregion

        #region IEnumerator Members
        int _ix = -1;
        object IEnumerator.Current
        {
            get 
            {
                return _cells[_cellIndex[_ix].ListPointer];
            }
        }

        bool IEnumerator.MoveNext()
        {
           _ix++;
           return _ix < _cells.Count;
        }

        void IEnumerator.Reset()
        {
            _ix = -1;
        }

        #endregion

        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this;
        }

        #endregion
    }
}
