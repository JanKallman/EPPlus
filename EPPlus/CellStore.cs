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
 * Jan Källman		    Added       		        2012-11-25
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using OfficeOpenXml;
    internal class IndexBase : IComparable<IndexBase>
    {
        internal short Index;
        public int CompareTo(IndexBase other)
        {
            return Index < other.Index ? -1 : Index > other.Index ? 1 : 0;
        }
    }
    internal class IndexItem :  IndexBase
    {
        internal int IndexPointer 
        {
            get; 
            set;
        }
    }
    internal class ColumnIndex : IndexBase, IDisposable
    {
        public ColumnIndex ()
	    {
            Pages=new PageIndex[32];
            PageCount=0;
	    }
        internal int GetPosition(int Row)
        {
            var page = (short)(Row >> CellStore<int>.pageBits);
            var res = Array.BinarySearch(Pages, 0, PageCount, new IndexBase() { Index = (short)(page) });
            if (res >= 0)
            {
                if (res+1 < PageCount && Pages[res+1].MinIndex <= Row)
                {
                    return res + 1;
                }
                else if (res - 1 >= 0 && Pages[res - 1].MaxIndex >= Row)
                {
                    return res - 1;
                }
                else
                {
                    return res;
                }
            }
            else
            {
                var p = ~res;
                if (res < 0  && p > 0 && p-1 < PageCount && PageCount > 0 && 
                    Pages[p-1].MaxIndex >= Row)
                {
                    return p - 1;
                }
                return res;
            }
        }
        internal int GetNextRow(int row)
        {
            //var page = (int)((ulong)row >> CellStore<int>.pageBits);
            var p = GetPosition(row);
            if (p < 0)
            {
                p = ~p;
                if (p >= PageCount)
                {
                    return -1;
                }
                else
                {
                    return Pages[p].IndexOffset + Pages[p].Rows[0].Index;
                }
            }
            else
            {
                var r=Pages[p].GetNextRow(row);
                if (r >= 0)
                {
                    return Pages[p].IndexOffset + Pages[p].Rows[r].Index;
                }
                else
                {
                    if (++p < PageCount)
                    {
                        return Pages[p].IndexOffset + Pages[p].Rows[0].Index;
                    }
                    else
                    {
                        return -1;
                    }
                }
            }
        }
        internal int FindNext(int Page)
        {
            var p = GetPosition(Page);
            if (p < 0)
            {
                return ~p;
            }
            return p;
        }
        internal PageIndex[] Pages;
        internal int PageCount;

        public void Dispose()
        {
            Pages = null;
        }
    }
    internal class PageIndex : IndexBase, IDisposable
    {
        public PageIndex ()
	    {
            Rows=new IndexItem[256];
            RowCount = 0;
	    }
        public PageIndex(IndexItem[] rows, int count)
        {
            Rows = rows;
            RowCount = count;
        }
        public PageIndex(PageIndex pageItem, int start, int size)
            :this(pageItem, start, size, pageItem.Index, pageItem.Offset)
        {

        }
        public PageIndex(PageIndex pageItem, int start, int size, short index, int offset)
        {
            Rows = new IndexItem[CellStore<int>.GetSize(size)];
            Array.Copy(pageItem.Rows, start, Rows,0, pageItem.RowCount);
            RowCount = size;
            Index = index;
            Offset = offset;
        }
        internal int Offset = 0;
        internal int IndexOffset
        {
            get
            {
                return (Index << CellStore<int>.pageBits) + (int)Offset;
            }
        }
        internal IndexItem[] Rows{ get; set;}
        internal int RowCount;

        internal int GetPosition(int offset)
        {
            return Array.BinarySearch(Rows, 0, RowCount, new IndexBase() { Index = (short)(offset) });
        }
        internal int GetNextRow(int row)
        {
            int offset = row - IndexOffset;
            var o= GetPosition(offset);
            if (o < 0)
            {
                o = ~o;
                if (o < RowCount)
                {
                    return o;
                }
                else
                {
                    return -1;
                }
            }
            return o;
        }

        public int MinIndex
        {
            get
            {
                if (Rows.Length > 0)
                {
                    return IndexOffset + Rows[0].Index;
                }
                else
                {
                    return -1;
                }
            }
        }
        public int MaxIndex
        {
            get
            {
                if (RowCount > 0)
                {
                    return IndexOffset + Rows[RowCount-1].Index;
                }
                else
                {
                    return -1;
                }
            }
        }

        public void Dispose()
        {
            Rows = null;
        }
    }
    /// <summary>
    /// This is the store for all Rows, Columns and Cells.
    /// It is a Dictionary implementation that allows you to change the Key (the RowID, ColumnID or CellID )
    /// </summary>
    internal class CellStore<T> : IDisposable// : IEnumerable<ulong>, IEnumerator<ulong>
    {
        internal const int pageBits = 13;   //8096
        internal const int PageSize = 1 << pageBits;
        internal const int PageSizeMax = PageSize << 1; //Double page size

        List<T> _values = new List<T>();
        internal ColumnIndex[] _columnIndex;
        internal int ColumnCount;
        public CellStore ()
	    {
            _columnIndex=new ColumnIndex[32];
	    }
        internal int GetPosition(int Column)
        {
            return Array.BinarySearch(_columnIndex, 0, ColumnCount, new IndexBase() { Index = (short)(Column) });
        }
        internal CellStore<T> Clone()
        {
            int row,col;
            var ret=new CellStore<T>();
            for (int c = 0; c < ColumnCount; c++)
            {
                col = _columnIndex[c].Index;
                for (int p = 0;p < _columnIndex[c].PageCount; p++)
                {
                    for (int r = 0; r < _columnIndex[c].Pages[p].RowCount; r++)
                    {
                        row = _columnIndex[c].Pages[p].IndexOffset + _columnIndex[c].Pages[p].Rows[r].Index;
                        ret.SetValue(row, col, _values[_columnIndex[c].Pages[p].Rows[r].IndexPointer]);
                    }
                }
            }
            return ret;
        }
        internal int Count
        {
            get
            {
                int count=0;
                for (int c = 0; c < ColumnCount; c++)
                {
                    for (int p = 0; p < _columnIndex[c].PageCount; p++)
                    {
                        count += _columnIndex[c].Pages[p].RowCount;
                    }
                }
                return count;
            }
        }
        internal bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
        {
            if (ColumnCount == 0)
            {
                fromRow = fromCol = toRow = toCol = 0;
                return false;
            }
            else
            {
                fromCol=_columnIndex[0].Index;
                if (fromCol <= 0 && ColumnCount > 1)
                {
                    fromCol = _columnIndex[1].Index;
                }
                else if(ColumnCount == 1 && fromCol <= 0)
                {
                    fromRow = fromCol = toRow = toCol = 0;
                    return false;
                }
                toCol=_columnIndex[ColumnCount-1].Index;
                fromRow = toRow= 0;

                for (int c = 0; c < ColumnCount; c++)
                {                    
                    int first, last;
                    if (_columnIndex[c].Pages[0].Rows[0].Index > 0)
                    {
                        first = _columnIndex[c].Pages[0].IndexOffset + _columnIndex[c].Pages[0].Rows[0].Index;
                    }
                    else
                    {
                        if(_columnIndex[c].Pages[0].RowCount>1)
                        {
                            first = _columnIndex[c].Pages[0].IndexOffset + _columnIndex[c].Pages[0].Rows[1].Index;
                        }
                        else if (_columnIndex[c].PageCount > 1)
                        {
                            first = _columnIndex[c].Pages[0].IndexOffset + _columnIndex[c].Pages[1].Rows[0].Index;
                        }
                        else
                        {
                            first = 0;
                        }
                    }
                    var lp = _columnIndex[c].PageCount - 1;
                    while(_columnIndex[c].Pages[lp].RowCount==0 && lp!=0)
                    {
                        lp--;
                    }
                    var p = _columnIndex[c].Pages[lp];
                    last = p.IndexOffset + p.Rows[lp].Index;

                    if (first > 0 && (first < fromRow || fromRow == 0))
                    {
                        fromRow=first;
                    }
                    if (first>0 && (last > toRow || toRow == 0))
                    {
                        toRow=last;
                    }
                }
                if (fromRow <= 0 || toRow <= 0)
                {
                    fromRow = fromCol = toRow = toCol = 0;
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
        internal int FindNext(int Column)
        {
            var c = GetPosition(Column);
            if (c < 0)
            {
                return ~c;
            }
            return c;
        }
        internal T GetValue(int Row, int Column)
        {
            int i = GetPointer(Row, Column);
            if (i >= 0)
            {
                return _values[i];
            }
            else
            {
                return default(T);
            }
            //var col = GetPosition(Column);
            //if (col >= 0)  
            //{
            //    var pos = _columnIndex[col].GetPosition(Row);
            //    if (pos >= 0) 
            //    {
            //        var pageItem = _columnIndex[col].Pages[pos];
            //        if (pageItem.MinIndex > Row)
            //        {
            //            pos--;
            //            if (pos < 0)
            //            {
            //                return default(T);
            //            }
            //            else
            //            {
            //                pageItem = _columnIndex[col].Pages[pos];
            //            }
            //        }
            //        short ix = (short)(Row - pageItem.IndexOffset);
            //        var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, new IndexBase() { Index = ix });
            //        if (cellPos >= 0) 
            //        {
            //            return _values[pageItem.Rows[cellPos].IndexPointer];
            //        }
            //        else //Cell does not exist
            //        {
            //            return default(T);
            //        }
            //    }
            //    else //Page does not exist
            //    {
            //        return default(T);
            //    }
            //}
            //else //Column does not exist
            //{
            //    return default(T);
            //}
        }
        int GetPointer(int Row, int Column)
        {
            var col = GetPosition(Column);
            if (col >= 0)
            {
                var pos = _columnIndex[col].GetPosition(Row);
                if (pos >= 0)
                {
                    var pageItem = _columnIndex[col].Pages[pos];
                    if (pageItem.MinIndex > Row)
                    {
                        pos--;
                        if (pos < 0)
                        {
                            return -1;
                        }
                        else
                        {
                            pageItem = _columnIndex[col].Pages[pos];
                        }
                    }
                    short ix = (short)(Row - pageItem.IndexOffset);
                    var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, new IndexBase() { Index = ix });
                    if (cellPos >= 0)
                    {
                        return pageItem.Rows[cellPos].IndexPointer;
                    }
                    else //Cell does not exist
                    {
                        return -1;
                    }
                }
                else //Page does not exist
                {
                    return -1;
                }
            }
            else //Column does not exist
            {
                return -1;
            }
        }
        internal bool Exists(int Row,int Column)
        {
            return GetPointer(Row, Column)>=0;
        }
        internal bool Exists(int Row, int Column, ref T value)
        {
            var p=GetPointer(Row, Column);
            if (p >= 0)
            {
                value = _values[p];
                return true;
            }
            else
            {                
                return false;
            }
        }
        internal void SetValue(int Row, int Column, T Value)
        {
            var col = Array.BinarySearch(_columnIndex, 0, ColumnCount, new IndexBase() { Index = (short)(Column) });
            var page = (short)(Row >> pageBits);
            if (col >= 0)
            {
                //var pos = Array.BinarySearch(_columnIndex[col].Pages, 0, _columnIndex[col].Count, new IndexBase() { Index = page });
                var pos = _columnIndex[col].GetPosition(Row);
                if(pos < 0)
                {
                    pos = ~pos;
                    if (pos - 1 < 0 || _columnIndex[col].Pages[pos - 1].IndexOffset + PageSize <= Row)
                    {
                        AddPage(_columnIndex[col], pos, page);
                    }
                    else
                    {
                        pos--;
                    }
                }
                var pageItem = _columnIndex[col].Pages[pos];
                if (pageItem.IndexOffset > Row)
                {
                    pos--;
                    page--;
                    if (pos < 0)
                    {
                        throw(new Exception("Unexpected error when setting value"));
                    }
                    pageItem = _columnIndex[col].Pages[pos];
                }

                short ix = (short)(Row - ((pos << pageBits) + pageItem.Offset));
                var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, new IndexBase() { Index = ix });
                if (cellPos < 0)
                {
                    cellPos = ~cellPos;
                    AddCell(pageItem, cellPos, ix, Value);
                }
                else
                {
                    _values[pageItem.Rows[cellPos].IndexPointer] = Value;
                }
            }
            else //Column does not exist
            {
                col = ~col;
                AddColumn(col, Column);
                AddPage(_columnIndex[col], 0, page);
                short ix = (short)(Row - (page << pageBits));
                AddCell(_columnIndex[col].Pages[0], 0, ix, Value);
            }
        }
        internal void Insert(int fromRow, int fromCol, int rows, int columns)
        {
            if (columns > 0)
            {
                var col = GetPosition(fromCol);
                if (col < 0)
                {
                    col = ~col;
                }
                for (var c = col; c < ColumnCount; c++)
                {
                    _columnIndex[c].Index += (short)columns;
                }
            }
            else
            {
                var page = fromRow >> pageBits;
                for(int c=0;c< ColumnCount;c++)
                {
                    var column = _columnIndex[c];
                    var pagePos = column.GetPosition(fromRow);
                    if (pagePos >= 0)
                    {
                        if (fromRow-1 >= column.Pages[pagePos].IndexOffset && fromRow-1<=column.Pages[pagePos].MaxIndex) //The row is inside the page
                        {
                           int offset = fromRow - column.Pages[pagePos].IndexOffset;
                           var rowPos = column.Pages[pagePos].GetPosition(offset);
                           if (rowPos < 0) 
                           {
                               rowPos = ~rowPos;
                           }
                           UpdateIndexOffset(column, pagePos, rowPos, fromRow, rows);
                        }
                        else if (column.Pages[pagePos].MinIndex > fromRow-1) //The row is on the page before.
                        {
                            int offset = fromRow - ((page - 1) << pageBits);
                            var rowPos = column.Pages[pagePos-1].GetPosition(offset);
                            if (rowPos > 0 && pagePos>0)
                            {
                                UpdateIndexOffset(column, pagePos-1, rowPos, fromRow, rows);
                            }
                        }
                        else
                        {

                        }
                    }
                    else
                    {
                        UpdateIndexOffset(column, ~pagePos, 0, fromRow, rows);
                    }
                }
            }
        }
        internal void Clear(int fromRow, int fromCol, int rows, int columns)
        {
            Delete(fromRow, fromCol, rows, columns, false);
        }
        internal void Delete(int fromRow, int fromCol, int rows, int columns)
        {
            Delete(fromRow, fromCol, rows, columns, true);
        }
        internal void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
        {
            if (columns > 0 && fromRow==1 && rows>=ExcelPackage.MaxRows)
            {
                DeleteColumns(fromCol, columns, shift);
            }
            else
            {
                var toCol = fromCol + columns - 1;
                var pageFromRow = fromRow >> pageBits;
                for (int c = fromCol; c < ColumnCount; c++)
                {
                    var column = _columnIndex[c];
                    if (column.Index > toCol) break;
                    var pagePos = column.GetPosition(fromRow);
                    if (pagePos < 0) pagePos = ~pagePos;
                    var page = column.Pages[pagePos];

                    if (page.RowCount > 0 && page.IndexOffset <= fromRow && page.IndexOffset + page.Rows[page.RowCount - 1].Index >= fromRow) //The row is inside the page
                    {
                        var endRow = fromRow+rows;
                        var delEndRow=DeleteCells(column.Pages[pagePos], fromRow, rows);
                        if (endRow > delEndRow && column.Pages[pagePos].MaxIndex < endRow && pagePos < column.PageCount - 1)
                        {
                            delEndRow = DeletePage(delEndRow, endRow - delEndRow, column, pagePos + 1);
                            if (shift) UpdatePageOffset(column, pagePos, endRow - delEndRow);
                            if (endRow > delEndRow)
                            {
                                pagePos = column.GetPosition(fromRow);
                                delEndRow = DeleteCells(column.Pages[pagePos], fromRow, endRow - delEndRow);
                                if (shift) UpdatePageOffset(column, pagePos, endRow - delEndRow);
                            }
                        }
                        else
                        {
                            if (shift) UpdatePageOffset(column, pagePos, rows);
                        }
                    }
                    else if (column.Pages[pagePos].IndexOffset > fromRow) //The row is on the page before.
                    {
                        int offset = fromRow + rows - 1 - ((pageFromRow - 1) << pageBits);
                        var rowPos = column.Pages[pagePos - 1].GetPosition(offset);
                        if (rowPos > 0 && pagePos > 0)
                        {
                            if (shift) UpdateIndexOffset(column, pagePos - 1, rowPos, fromRow + rows - 1, -rows);
                        }
                    }
                    else
                    {

                    }
                }
            }
        }
        private void UpdatePageOffset(ColumnIndex column, int pagePos, int rows)
        {
            //Update Pageoffset
            
            if (++pagePos < column.PageCount)
            {
                for (int p = pagePos; p < column.PageCount; p++)
                {
                    column.Pages[p].Offset -= rows;
                }

                if (Math.Abs(column.Pages[pagePos].Offset) > PageSize ||
                    Math.Abs(column.Pages[pagePos].Rows[column.Pages[pagePos].RowCount-1].Index) > PageSizeMax) //Split or Merge???
                {
                    rows=ResetPageOffset(column, pagePos, rows);
                    ////MergePages
                    //if (column.Pages[pagePos - 1].Index + 1 == column.Pages[pagePos].Index)
                    //{
                    //    if (column.Pages[pagePos].IndexOffset + column.Pages[pagePos].Rows[column.Pages[pagePos].RowCount - 1].Index + rows -
                    //        column.Pages[pagePos - 1].IndexOffset + column.Pages[pagePos - 1].Rows[0].Index <= PageSize)
                    //    {
                    //        //Merge
                    //        MergePage(column, pagePos - 1, -rows);
                    //    }
                    //    else
                    //    {
                    //        //Split
                    //    }
                    //}
                    //rows -= PageSize;
                    //for (int p = pagePos; p < column.PageCount; p++)
                    //{                            
                    //    column.Pages[p].Index -= 1;
                    //}
                    return;
                }
            }
        }

        private int ResetPageOffset(ColumnIndex column, int pagePos, int rows)
        {
            PageIndex fromPage=column.Pages[pagePos];
            PageIndex toPage;
            short pageAdd = 0;
            if (fromPage.Offset < -PageSize)
            {
                toPage=column.Pages[pagePos-1];
                pageAdd = -1;
                if (fromPage.Index - 1 == toPage.Index)
                {
                    if (fromPage.IndexOffset + fromPage.Rows[fromPage.RowCount - 1].Index -
                        toPage.IndexOffset + toPage.Rows[0].Index <= PageSizeMax)
                    {
                        MergePage(column, pagePos - 1);
                        //var newPage = new PageIndex(toPage, 0, GetSize(fromPage.RowCount + toPage.RowCount));
                        //newPage.RowCount = fromPage.RowCount + fromPage.RowCount;
                        //Array.Copy(toPage.Rows, 0, newPage.Rows, 0, toPage.RowCount);
                        //Array.Copy(fromPage.Rows, 0, newPage.Rows, toPage.RowCount, fromPage.RowCount);
                        //for (int r = toPage.RowCount; r < newPage.RowCount; r++)
                        //{
                        //    newPage.Rows[r].Index += (short)(fromPage.IndexOffset - toPage.IndexOffset);
                        //}
                        
                    }
                }
                else //No page after 
                {
                    fromPage.Index -= pageAdd;
                    fromPage.Offset += PageSize;
                }
            }
            else if (fromPage.Offset > PageSize)
            {
                toPage = column.Pages[pagePos + 1];
                pageAdd = 1;
                if (fromPage.Index + 1 == toPage.Index)
                {

                }
                else
                {
                    fromPage.Index += pageAdd;
                    fromPage.Offset += PageSize;
                }
            }
            return rows;
        }

        private int DeletePage(int fromRow, int rows, ColumnIndex column, int pagePos)
        {
            PageIndex page = column.Pages[pagePos];
            while (page != null && page.IndexOffset >= fromRow && page.IndexOffset + page.Rows[page.RowCount - 1].Index <= fromRow + rows)
            {
                //Delete entire page.
                var pageEndRow = page.Rows[page.RowCount - 1].Index;
                var prevOffset = page.Offset;
                Array.Copy(column.Pages, pagePos + 1, column.Pages, pagePos, column.PageCount - pagePos + 1);
                column.PageCount--;
                for (int i = pagePos; i < column.PageCount; i++)
                {
                    column.Pages[i].Index--;
                }
                if (column.PageCount > pagePos)
                {
                    page = column.Pages[pagePos];
                    page.Offset = prevOffset;
                }
                else
                {
                    //No more pages, return 0
                    return 0;
                }
                rows -= pageEndRow;
            }
            return rows;
        }

        private int DeleteCells(PageIndex page,  int fromRow, int toRow)
        {
            var fromPos = page.GetPosition(fromRow - (page.IndexOffset));
            if (fromPos < 0)
            {
                fromPos = ~fromPos;
            }
            var maxRow = page.MaxIndex;
            var offset = toRow - page.IndexOffset;
            if (offset > PageSizeMax) offset = PageSizeMax;
            var toPos = page.GetPosition(offset);
            if (toPos < 0)
            {
                toPos = ~toPos;
            }
            
            if (fromPos <= toPos && fromPos < page.RowCount && page.Rows[fromPos].Index < toRow - 1)
            {
                if (toPos >= page.RowCount-1)
                {
                    var deletedRow = page.IndexOffset + page.Rows[page.RowCount - 1].Index - fromRow;
                    page.RowCount -= deletedRow;
                    return deletedRow;
                }
                else
                {
                    int rows = toRow - fromRow;
                    for (int r = toPos; r < page.RowCount; r++)
                    {
                        page.Rows[r].Index -= (short)rows;
                    }
                    Array.Copy(page.Rows, toPos, page.Rows, fromPos, page.RowCount - toPos);
                    page.RowCount -= rows;

                    return toRow;
                }
            }
            return toRow < maxRow ? toRow : maxRow;
        }

        private void DeleteColumns(int fromCol, int columns, bool shift)
        {
            var fPos = GetPosition(fromCol);
            if (fPos < 0)
            {
                fPos = ~fPos;
            }
            int tPos = fPos;
            for (var c = fPos; c < ColumnCount; c++)
            {
                if (_columnIndex[c].Index < fromCol + columns) break;
                tPos = c;
            }

            if (_columnIndex[fPos].Index >= fromCol && _columnIndex[fPos].Index <= fromCol + columns)
            {
                if (_columnIndex[fPos].Index > ColumnCount)
                {
                    Array.Copy(_columnIndex, fPos, _columnIndex, tPos, tPos - fPos);
                }
                ColumnCount -= columns;
            }
            if (shift)
            {
                for (var c = tPos + 1; c < ColumnCount; c++)
                {
                    _columnIndex[c].Index -= (short)columns;
                }
            }
        }

        private void UpdateIndexOffset(ColumnIndex column, int pagePos, int rowPos, int row, int rows)
        {
            if (pagePos >= column.PageCount) return;    //A page after last cell.
            var page = column.Pages[pagePos];
            if (rows > PageSize)
            {
                short addPages = (short)(rows >> pageBits);
                int offset = +(int)(rows - (PageSize*addPages));
                for (int p = pagePos + 1; p < column.PageCount; p++)
                {
                    if (column.Pages[p].Offset + offset > PageSize)
                    {
                        column.Pages[p].Index += (short)(addPages + 1);
                        column.Pages[p].Offset += PageSize - offset;
                    }
                    else
                    {
                        column.Pages[p].Index += addPages;
                        column.Pages[p].Offset += offset;
                    }
                    
                }

                var size = page.RowCount - rowPos;
                if (page.RowCount > rowPos)
                {
                    if (column.PageCount == pagePos) //No page after, create a new one.
                    {
                        //Copy rows to next page.
                        var newPage = CopyNew(page, rowPos, size);
                        newPage.Index = (short)(pagePos + addPages + 1);
                        newPage.Offset = row + rows - (newPage.Index * PageSize);
                        AddPage(column, pagePos + 1, newPage);
                    }
                    else
                    {
                        if (column.Pages[pagePos + 1].RowCount + size > PageSizeMax) //Split Page
                        {
                            //MergePage(page, rowPos, rows, size, column.Pages[pagePos + 1]);
                        }
                        else //Copy Page.
                        {
                            CopyMergePage(page, rowPos, rows, size, column.Pages[pagePos + 1]);                            
                        }
                    }
                }
            }
            else
            {
                //Add to Pages.
                for (int r = rowPos; r < page.RowCount; r++)
                {
                    if (page.Rows[r].Index + rows > PageSizeMax)
                    {
                        IndexItem[] newRows = new IndexItem[GetSize(page.RowCount - page.Rows[r].Index)];
                        var newPage = new PageIndex(newRows, r);
                        newPage.Index = (short)(pagePos + 1);
                        //TODO: MoveRows to next page.
                    }
                    else
                    {
                        page.Rows[r].Index += (short)rows;
                    }
                }

                for (int p = pagePos + 1; p < column.PageCount; p++)
                {
                    if (column.Pages[p].Offset + rows < PageSizeMax)
                    {
                        column.Pages[p].Offset += rows;
                    }
                    else
                    {
                        //TODO:Split page 
                    }
                }
            }
        }

        private void CopyMergePage(PageIndex page, int rowPos, int rows, int size, PageIndex ToPage)
        {
            var startRow = page.IndexOffset + page.Rows[rowPos].Index + rows;
            var newRows = new IndexItem[GetSize(ToPage.Rows.Length + size)];
            page.RowCount -= size;
            Array.Copy(page.Rows, rowPos, newRows, 0, size);
            for (int r = 0; r < size; r++)
            {
                newRows[r].Index += (short)(page.IndexOffset + rows - ToPage.IndexOffset);
            }

            Array.Copy(ToPage.Rows, 0, newRows, size, ToPage.Rows.Length);
            ToPage.Rows = newRows;
            ToPage.RowCount += size;
        }
        private void MergePage(ColumnIndex column, int pagePos)
        {
            PageIndex Page1=column.Pages[pagePos];
            PageIndex Page2 = column.Pages[pagePos + 1];

            var newPage = new PageIndex(Page1, 0, Page1.RowCount + Page2.RowCount);
            newPage.RowCount = Page1.RowCount + Page2.RowCount;
            Array.Copy(Page1.Rows, 0, newPage.Rows, 0, Page1.RowCount);
            Array.Copy(Page2.Rows, 0, newPage.Rows, Page1.RowCount, Page2.RowCount);
            for (int r = Page1.RowCount; r < newPage.RowCount; r++)
            {
                newPage.Rows[r].Index += (short)(Page2.IndexOffset - Page1.IndexOffset);
            }

            column.Pages[pagePos] = newPage;
            column.PageCount--;

            if (column.PageCount > (pagePos + 1))
            {
                Array.Copy(column.Pages, pagePos+2, column.Pages,pagePos+1,column.PageCount-(pagePos+1));
                for (int p = pagePos + 1; p < column.PageCount; p++)
                {
                    column.Pages[p].Index--;
                    column.Pages[p].Offset += PageSize;
                }
            }
        }

        private PageIndex CopyNew(PageIndex pageFrom, int rowPos, int size)
        {
            IndexItem[] newRows = new IndexItem[GetSize(size)];
            Array.Copy(pageFrom.Rows, rowPos, newRows, 0, size);
            return new PageIndex(newRows, size);
        }

        internal static int GetSize(int size)
        {
            var newSize=256;
            while (newSize < size)
            {
                newSize <<= 1;
            }
            return newSize;
        }
        private void AddCell(PageIndex pageItem, int pos, short ix, T value)
        {
            if (pageItem.RowCount == pageItem.Rows.Length)
            {
                if (pageItem.RowCount == PageSizeMax) //Max size-->Split
                {
                    var newPage = new PageIndex(pageItem, 0, PageSize);
                    var nextPage = new PageIndex(pageItem, PageSize, PageSize, (short)(pageItem.Index+1), pageItem.Offset);
                    for(int r=0;r<nextPage.RowCount;r++)
                    {
                        nextPage.Rows[r].Index=(short)(nextPage.Rows[r].Index-PageSize);
                    }
                }
                else //Expand to double size.
                {
                    var rowsTmp = new IndexItem[pageItem.Rows.Length << 1];
                    Array.Copy(pageItem.Rows, 0, rowsTmp, 0, pageItem.RowCount);
                    pageItem.Rows = rowsTmp;
                }
            }
            if (pos < pageItem.RowCount)
            {
                Array.Copy(pageItem.Rows, pos, pageItem.Rows, pos + 1, pageItem.RowCount - pos);
            }
            pageItem.Rows[pos] = new IndexItem() { Index = ix,IndexPointer=_values.Count };
            _values.Add(value);
            pageItem.RowCount++;
        }
        private void AddPage(ColumnIndex column, int pos, short index)
        {
            AddPage(column, pos);
            column.Pages[pos] = new PageIndex() { Index = index };
            if (pos > 0)
            {
                var pp=column.Pages[pos-1];
                if(pp.RowCount>0 && pp.Rows[pp.RowCount-1].Index > PageSize)
                {
                    column.Pages[pos].Offset = pp.Rows[pp.RowCount-1].Index-PageSize;
                }
            }
        }
        /// <summary>
        /// Add a new page to the collection
        /// </summary>
        /// <param name="column">The column</param>
        /// <param name="pos">Position</param>
        /// <param name="page">The new page object to add</param>
        private void AddPage(ColumnIndex column, int pos, PageIndex page)
        {
            AddPage(column, pos);
            column.Pages[pos] = page ;
        }
        /// <summary>
        /// Add a new page to the collection
        /// </summary>
        /// <param name="column">The column</param>
        /// <param name="pos">Position</param>
        private void AddPage(ColumnIndex column, int pos)
        {
            if (column.PageCount ==column.Pages.Length)
            {
                var pageTmp = new PageIndex[column.Pages.Length * 2];
                Array.Copy(column.Pages, 0, pageTmp, 0, column.PageCount);
                column.Pages = pageTmp;
            }
            if (pos < column.PageCount)
            {
                Array.Copy(column.Pages, pos, column.Pages, pos + 1, column.PageCount - pos);
            }
            column.PageCount++;
        }
        private void AddColumn(int pos, int Column)
        {
            if (ColumnCount == _columnIndex.Length)
            {
                var colTmp = new ColumnIndex[_columnIndex.Length*2];
                Array.Copy(_columnIndex, 0, colTmp, 0, ColumnCount);
                _columnIndex = colTmp;
            }
            if (pos < ColumnCount)
            {
                Array.Copy(_columnIndex, pos, _columnIndex, pos + 1, ColumnCount - pos);
            }
            _columnIndex[pos] = new ColumnIndex() { Index = (short)(Column) };
            ColumnCount++;
        }        
        int _colPos = -1, _row, _col;
        public ulong Current
        {
            get
            {
                return ((ulong)_row << 32) | (uint)(_columnIndex[_colPos].Index);
            }
        }

        public void Dispose()
        {
            _values = null;
            _columnIndex = null;
        }

        //object IEnumerator.Current
        //{
        //    get 
        //    {
        //        return GetValue(_row+1, _columnIndex[_colPos].Index);
        //    }
        //}
        public bool MoveNext()
        {
            return GetNextCell(ref _row, ref _colPos, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        internal bool NextCell(ref int row, ref int col)
        {
            
            return NextCell(ref row, ref col, 0,0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        internal bool NextCell(ref int row, ref int col, int minRow, int minColPos,int maxRow, int maxColPos)
        {
            if (minColPos >= ColumnCount)
            {
                return false;
            }
            if (maxColPos >= ColumnCount)
            {
                maxColPos = ColumnCount-1;
            }
            var c=GetPosition(col);
            if(c>=0)
            {
                if (c > maxColPos)
                {
                    if (col <= minColPos)
                    {
                        return false;
                    }
                    col = minColPos;
                    return NextCell(ref row, ref col);
                }
                else
                {
                    return GetNextCell(ref row, ref c, minColPos, maxRow, maxColPos);
                }
            }
            else
            {
                c=~c;
                if (c > _columnIndex[c].Index)
                {
                    if (col <= minColPos)
                    {
                        return false;
                    }
                    col = minColPos;
                    return NextCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
                }
                else
                {
                    return GetNextCell(ref c, ref row, minColPos, maxRow, maxColPos);
                }
            }
        }
        internal bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos)
        {
            if (ColumnCount == 0)
            {
                return false;
            }
            else
            {
                if (++colPos < ColumnCount && colPos <=endColPos)
                {
                    var r = _columnIndex[colPos].GetNextRow(row);
                    if (r == row) //Exists next Row
                    {
                        return true;
                    }
                    else
                    {
                        int minRow, minCol;
                        if (r > row)
                        {
                            minRow = r;
                            minCol = colPos;
                        }
                        else
                        {
                            minRow = int.MaxValue;
                            minCol = 0;
                        }

                        var c = colPos + 1;
                        while (c < ColumnCount && c <= endColPos)
                        {
                            r = _columnIndex[c].GetNextRow(row);
                            if (r == row) //Exists next Row
                            {
                                colPos = c;
                                return true;
                            }
                            if (r > row && r < minRow)
                            {
                                minRow = r;
                                minCol = c;
                            }
                            c++;
                        }
                        c = startColPos;
                        if (row < endRow)
                        {
                            row++;
                            while (c < colPos)
                            {
                                r = _columnIndex[c].GetNextRow(row);
                                if (r == row) //Exists next Row
                                {
                                    colPos = c;
                                    return true;
                                }
                                if (r > row && (r < minRow || (r==minRow && c<minCol)) && r <= endRow)
                                {
                                    minRow = r;
                                    minCol = c;
                                }
                                c++;
                            }
                        }

                        if (minRow == int.MaxValue || minRow > endRow)
                        {
                            return false;
                        }
                        else
                        {
                            row = minRow;
                            colPos = minCol;
                            return true;
                        }
                    }
                }
                else
                {
                    if (colPos <= startColPos || row>=endRow)
                    {
                        return false;
                    }
                    colPos = startColPos - 1;
                    row++;
                    return GetNextCell(ref row, ref colPos, startColPos, endRow, endColPos);
                }
            }
        }
        internal bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos, ref int[] pagePos, ref int[] cellPos)
        {
            if (colPos == endColPos)
            {
                colPos = startColPos;
                row++;
            }
            else
            {
                colPos++;
            }

            if (pagePos[colPos] < 0)
            {
                if(pagePos[colPos]==-1)
                {
                    pagePos[colPos] = _columnIndex[colPos].GetPosition(row);
                }
            }
            else if (_columnIndex[colPos].Pages[pagePos[colPos]].RowCount <= row)
            {
                if (_columnIndex[colPos].PageCount > pagePos[colPos])
                    pagePos[colPos]++;
                else
                {
                    pagePos[colPos]=-2;
                }
            }
            
            var r = _columnIndex[colPos].Pages[pagePos[colPos]].IndexOffset + _columnIndex[colPos].Pages[pagePos[colPos]].Rows[cellPos[colPos]].Index;
            if (r == row)
            {
                row = r;
            }
            else
            {
            }
            return true;
        }
        internal bool PrevCell(ref int row, ref int col)
        {
            return PrevCell(ref row, ref col, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
        }
        internal bool PrevCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
        {
            if (minColPos >= ColumnCount)
            {
                return false;
            }
            if (maxColPos >= ColumnCount)
            {
                maxColPos = ColumnCount - 1;
            }
            var c = GetPosition(col);
            if(c>=0)
            {
                if (c == 0)
                {
                    if (col >= maxColPos)
                    {
                        return false;
                    }
                    col = maxColPos;
                    return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
                }
                else
                {
                    var ret=GetPrevCell(ref row, ref c, minRow, minColPos, maxColPos);
                    if (ret)
                    {
                        col = _columnIndex[c].Index;
                    }
                    return ret;
                }
            }
            else
            {
                c=~c;
                if (c == 0)
                {
                    if (col >= maxColPos)
                    {
                        return false;
                    }
                    col = maxColPos;
                    return PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
                }
                else
                {
                    var ret = GetPrevCell(ref row, ref c, minRow, minColPos, maxColPos);
                    if (ret)
                    {
                        col = _columnIndex[c].Index;
                    }
                    return ret;
                }
            }
        }
        internal bool GetPrevCell(ref int row, ref int colPos, int startRow, int startColPos, int endColPos)
        {
            if (ColumnCount == 0)
            {
                return false;
            }
            else
            {
                if (--colPos >= startColPos)
//                if (++colPos < ColumnCount && colPos <= endColPos)
                {
                    var r = _columnIndex[colPos].GetNextRow(row);
                    if (r == row) //Exists next Row
                    {
                        return true;
                    }
                    else
                    {
                        int minRow, minCol;
                        if (r > row && r >= startRow)
                        {
                            minRow = r;
                            minCol = colPos;
                        }
                        else
                        {
                            minRow = int.MaxValue;
                            minCol = 0;
                        }

                        var c = colPos + 1;
                        if (c <= endColPos)
                        {
                            while (c >= 0)
                            {
                                r = _columnIndex[c].GetNextRow(row);
                                if (r == row) //Exists next Row
                                {
                                    colPos = c;
                                    return true;
                                }
                                if (r > row && r < minRow && r >= startRow)
                                {
                                    minRow = r;
                                    minCol = c;
                                }
                                c--;
                            }
                        }
                        if (row > startRow)
                        {
                            c = endColPos;
                            row--;
                            while (c > colPos)
                            {
                                r = _columnIndex[c].GetNextRow(row);
                                if (r == row) //Exists next Row
                                {
                                    colPos = c;
                                    return true;
                                }
                                if (r > row && r < minRow && r >= startRow)
                                {
                                    minRow = r;
                                    minCol = c;
                                }
                                c--;
                            }
                        }
                        if (minRow == int.MaxValue || startRow < minRow)
                        {
                            return false;
                        }
                        else
                        {
                            row = minRow;
                            colPos = minCol;
                            return true;
                        }
                    }
                }
                else
                {
                    colPos = ColumnCount;
                    row--;
                    return GetPrevCell(ref colPos, ref row, startRow, startColPos, endColPos);
                }
            }
        }
        public void Reset()
        {
            _colPos = -1;            
            _row= 0;
        }

        //public IEnumerator<ulong> GetEnumerator()
        //{
        //    this.Reset();
        //    return this;
        //}

        //IEnumerator IEnumerable.GetEnumerator()
        //{
        //    this.Reset();
        //    return this;
        //}


    }
    internal class CellsStoreEnumerator<T>
    {
        CellStore<T> _cellStore;
        int row, colPos;
        int[] pagePos, cellPos;
        int minRow, minColPos, maxRow, maxColPos;
        public CellsStoreEnumerator(CellStore<T> cellStore) :
            this(cellStore, 0,0,ExcelPackage.MaxRows, ExcelPackage.MaxColumns)        
        {
        }
        public CellsStoreEnumerator(CellStore<T> cellStore, int StartRow, int StartCol, int EndRow, int EndCol)
        {
            _cellStore = cellStore;
            minRow = StartRow;
            maxRow = EndRow;
            minColPos = cellStore.GetPosition(StartCol);
            if (minColPos < 0) minColPos = ~minColPos;
            maxColPos = cellStore.GetPosition(EndCol);
            if (maxColPos < 0) maxColPos = ~maxColPos-1;
            row = minRow;
            colPos = minColPos - 1;

            var cols = maxColPos - minColPos + 1;
            pagePos = new int[cols];
            cellPos = new int[cols];
            for (int i = 0; i < cols; i++)
            {
                pagePos[i] = -1;
                cellPos[i] = -1;
            }

        }
        internal int Row 
        {
            get
            {
                return row;
            }
        }
        internal int Column
        {
            get
            {
                return _cellStore._columnIndex[colPos].Index;
            }
        }
        internal T Value
        {
            get
            {
                return _cellStore.GetValue(row, Column);
            }
            set
            {
                _cellStore.SetValue(row, Column,value);
            }
        }
        internal bool Next()
        {
            //return _cellStore.GetNextCell(ref row, ref colPos, minColPos, maxRow, maxColPos);
            return _cellStore.GetNextCell(ref row, ref colPos, minColPos, maxRow, maxColPos);
        }
        internal bool Previous()
        {
            return _cellStore.GetPrevCell(ref row, ref colPos, minRow, minColPos, maxColPos);
        }

        public string CellAddress 
        {
            get
            {
                return ExcelAddressBase.GetAddress(Row, Column);
            }
        }
    }
    internal class FlagCellStore : CellStore<byte>
    {
        internal void SetFlagValue(int Row, int Col, bool value, CellFlags cellFlags)
        {
            SetValue(Row, Col, (byte)(GetValue(Row, Col) | ((byte)cellFlags)));
        }
        internal bool GetFlagValue(int Row, int Col, CellFlags cellFlags)
        {
            return !(((byte)cellFlags & GetValue(Row, Col)) == 0);
        }
    }
