/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Added this class		        2010-01-28
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using OfficeOpenXml.Style;
using System.Xml;

namespace OfficeOpenXml
{
    /// <summary>
    /// A range of cells 
    /// </summary>
    public class ExcelRangeBase : ExcelAddress, IExcelCell, IDisposable
    {
        protected ExcelWorksheet _worksheet;

        #region "Constructors"
        protected internal ExcelRangeBase(ExcelWorksheet xlWorksheet)
        {
            _worksheet = xlWorksheet;
            if (_worksheet.View.SelectedRange == "")
            {
                _address = "A1";
                return;
            }
            else
            {
                _address = _worksheet.View.SelectedRange;
            }
            GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
        }
        protected internal ExcelRangeBase(ExcelWorksheet xlWorksheet, string address)
        {
            _worksheet = xlWorksheet;
            _address = address;
            GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
        }   
        #endregion
        #region "Public Properties"
        /// <summary>
        /// The styleobject for the range.
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return _worksheet.Workbook.Styles.GetStyleObject(_worksheet.Cell(_fromRow, _fromCol).StyleID, _worksheet.PositionID, _address);
            }
        }
        /// <summary>
        /// The named style
        /// </summary>
        public string StyleName
        {
            get
            {
                return _worksheet.Cell(_fromRow, _fromCol).StyleName;
            }
            set
            {
                int styleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _worksheet.Cell(row, col).SetNewStyleName(value, styleID);
                    }
                }
            }
        }
        /// <summary>
        /// The style ID. 
        /// It is not recomended to use this one. Use Named styles as an alternative.
        /// If you do, make sure that you use the Style.UpdateXml() method to update any new styles added to the workbook.
        /// </summary>
        public int StyleID
        {
            get
            {
                return _worksheet.Cell(_fromRow, _fromCol).StyleID;
            }
            set
            {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _worksheet.Cell(row, col).StyleID = value;
                    }
                }
            }
        }
        /// <summary>
        /// Set the range to a specific value
        /// </summary>
        public object Value
        {
            get
            {
                return _worksheet.Cell(_fromRow, _fromCol).Value;
            }
            set
            {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _worksheet.Cell(row, col).Value = value;
                    }
                }
            }
        }
        /// <summary>
        /// Gets or sets a formula for a range.
        /// </summary>
        public string Formula
        {
            get
            {
                return _worksheet.Cell(_fromRow, _fromCol).Formula;
            }
            set
            {
                 if (string.IsNullOrEmpty(value))
                 {
                     _worksheet.Cell(_fromRow, _fromCol).Formula = string.Empty;
                     return;
                 }
                 
                if (value[0] == '=') value = value.Substring(1, value.Length - 1); // remove any starting equalsign.
                //If formula spans only one cell, set the formula property
                if (_fromRow == _toRow && _fromCol == _toCol)
                {
                    _worksheet.Cell(_fromRow, _fromCol).Formula = value;
                }
                else //Otherwise we use a shared formula.
                {
                    RemoveFormuls();
                    ExcelWorksheet.Formulas f = new ExcelWorksheet.Formulas();
                    f.Formula = value;
                    f.Index = _worksheet.GetMaxShareFunctionIndex();
                    f.Address = _address;
                    f.StartCol = _fromCol;
                    f.StartRow = _fromRow;

                    _worksheet._sharedFormulas.Add(f.Index, f);
                    _worksheet.Cell(_fromRow, _fromCol).SharedFormulaID = f.Index;
                    _worksheet.Cell(_fromRow, _fromCol).Formula = value;

                    for (int col = _fromCol; col <= _toCol; col++)
                    {
                        for (int row = _fromRow; row <= _toRow; row++)
                        {
                            _worksheet.Cell(row, col).SharedFormulaID = f.Index;
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Gets or Set a formula in R1C1 format.
        /// </summary>
        public string FormulaR1C1
        {
            get
            {
                return _worksheet.Cell(_fromRow, _fromCol).FormulaR1C1;
            }
            set
            {
                if (value[0] == '=') value = value.Substring(1, value.Length - 1); // remove any starting equalsign.
                Formula = ExcelCell.TranslateFromR1C1(value, _fromRow, _fromCol);
            }
        }
        /// <summary>
        /// Set the hyperlink property for a range of cells
        /// </summary>
        public Uri Hyperlink
        {
            get
            {
                return _worksheet.Cell(_fromRow, _fromCol).Hyperlink;
            }
            set
            {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _worksheet.Cell(row, col).Hyperlink = value;
                    }
                }
            }
        }
        /// <summary>
        /// If the cells in the range are merged.
        /// </summary>
        public bool Merge
        {
            get
            {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        if (!_worksheet.Cell(row, col).Merge)
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            set
            {
                if (!value)
                {
                    if (_worksheet.MergedCells.List.Contains(_address))
                    {
                        SetCellMerge(false);
                        _worksheet.MergedCells.List.Remove(_address);
                    }
                    else if (!CheckMergeDiff(false))
                    {
                        throw (new Exception("Range is not fully merged.Specify the exact range"));
                    }
                }
                else
                {
                    if (CheckMergeDiff(false))
                    {
                        SetCellMerge(true);
                        _worksheet.MergedCells.List.Add(_address);
                    }
                    else
                    {
                        if (!_worksheet.MergedCells.List.Contains(_address))
                        {
                            throw (new Exception("Cells are already merged"));
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Set an autofilter for the range
        /// </summary>
        public bool AutoFilter
        {
            get
            {
                ExcelAddressBase address = _worksheet.AutoFilterAddress;
                if (_fromRow >= address.Start.Row
                    &&
                    _toRow <= address.End.Row
                    &&
                    _fromCol >= address.Start.Column
                    &&
                    _toCol <= address.End.Column)
                {
                    return true;
                }
                return false;
            }
            set
            {
                _worksheet.AutoFilterAddress = new ExcelAddressBase(_address);
                if (_worksheet.Names.ContainsKey("_xlnm._FilterDatabase"))
                {
                    _worksheet.Names.Remove("_xlnm._FilterDatabase");
                }
                var result = _worksheet.Names.Add("_xlnm._FilterDatabase", this);
                result.IsNameHidden = true;
            }
        }        
        /// <summary>
        /// If the value is in richtext format.
        /// Then the value propery contains the raw XML. Please check the openXML documentation for info;
        /// </summary>
        public bool IsRichText
        {
            get
            {
                return _worksheet.Cell(_fromRow, _fromCol).IsRichText;
            }
            set
            {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _worksheet.Cell(row, col).IsRichText = value;
                    }
                }
            }
        }
        ExcelRichTextCollection _rtc = null;
        public ExcelRichTextCollection RichText
        {
            get
            {
                if (_rtc == null)
                {
                    XmlDocument xml = new XmlDocument();
                    if (_worksheet.Cell(_fromRow, _fromCol).Value != null)
                    {
                        xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" >" + _worksheet.Cell(_fromRow, _fromCol).Value.ToString() + "</si>");
                    }
                    else
                    {
                        xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />");
                    }
                    _rtc = new ExcelRichTextCollection(_worksheet.NameSpaceManager, xml.SelectSingleNode("d:si", _worksheet.NameSpaceManager), this);
                }
                return _rtc;
            }
        }
        public ExcelComment Comment
        {
            get
            {
                ulong cellID= GetCellID(_worksheet.SheetID, _fromRow, _fromCol);
                if(_worksheet._comments!=null && _worksheet._comments._comments.ContainsKey(cellID))
                {
                    return _worksheet._comments._comments[cellID] as ExcelComment;
                }
                return null;
            }
        }
        /// <summary>
        /// WorkSheet object 
        /// </summary>
        public ExcelWorksheet Worksheet
        {
            get
            {
                return _worksheet;
            }
        }
        /// <summary>
        /// Address including sheetname
        /// </summary>
        public string FullAddress
        {
            get
            {
                return GetFullAddress(_worksheet.Name, _address);
            }
        }
        /// <summary>
        /// Address including sheetname
        /// </summary>
        public string FullAddressAbsolute
        {
            get
            {
                return GetFullAddress(_worksheet.Name, GetAddress(_fromRow, _fromCol, _toRow, _toCol, true));
            }
        }
        //ExcelCellAddress _startPosition=null;
        //public ExcelCellAddress Start
        //{
        //    get
        //    {
        //        if (_startPosition == null)
        //        {
        //            _startPosition = new ExcelCellAddress(_fromRow, _fromCol);
        //        }
        //        return _startPosition;
        //    }
        //}
        //ExcelCellAddress _endPosition = null;
        ///// <summary>
        ///// Gets the row and column if the bottom right cell.
        ///// </summary>
        ///// <value>The end row and column.</value>
        //public ExcelCellAddress End
        //{
        //    get
        //    {
        //        if (_endPosition == null)
        //        {
        //            _endPosition = new ExcelCellAddress(_toRow, _toCol);
        //        }
        //        return _endPosition;
        //    }
        //}

        #endregion
        #region "Private Methods"
        /// <summary>
        /// Check if the range is partly merged
        /// </summary>
        /// <returns></returns>
        private bool CheckMergeDiff()
        {
            return CheckMergeDiff(_worksheet.Cell(_fromRow, _fromCol).Merge);
        }
        /// <summary>
        /// Check if the range is partly merged
        /// </summary>
        /// <param name="startValue">the starting value</param>
        /// <returns></returns>
        private bool CheckMergeDiff(bool startValue)
        {
            for (int col = _fromCol; col <= _toCol; col++)
            {
                for (int row = _fromRow; row <= _toRow; row++)
                {
                    if (_worksheet.Cell(row, col).Merge != startValue)
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        /// <summary>
        /// Set the merge flag for the range
        /// </summary>
        /// <param name="value"></param>
        internal void SetCellMerge(bool value)
        {
            for (int col = _fromCol; col <= _toCol; col++)
            {
                for (int row = _fromRow; row <= _toRow; row++)
                {
                    _worksheet.Cell(row, col).Merge = value;
                }
            }
        }
        /// <summary>
        /// Set the value without altering the richtext property
        /// </summary>
        /// <param name="value">the value</param>
        internal void SetValueRichText(object value)
        {
            for (int col = _fromCol; col <= _toCol; col++)
            {
                for (int row = _fromRow; row <= _toRow; row++)
                {
                    _worksheet.Cell(row, col).SetValueRichText(value);
                }
            }
        }
        /// <summary>
        /// Removes a shared formula
        /// </summary>
        private void RemoveFormuls()
        {
            List<int> removed = new List<int>();
            int fFromRow, fFromCol, fToRow, fToCol;
            foreach (int index in _worksheet._sharedFormulas.Keys)
            {
                ExcelWorksheet.Formulas f = _worksheet._sharedFormulas[index];
                ExcelCell.GetRowColFromAddress(f.Address, out fFromRow, out fFromCol, out fToRow, out fToCol);
                if (((fFromCol >= _fromCol && fFromCol <= _toCol) ||
                   (fToCol >= _fromCol && fToCol <= _toCol)) &&
                   ((fFromRow >= _fromRow && fFromRow <= _toRow) ||
                   (fToRow >= _fromRow && fToRow <= _toRow)))
                {
                    for (int col = fFromCol; col <= fToCol; col++)
                    {
                        for (int row = fFromRow; row <= fToRow; row++)
                        {
                            _worksheet.Cell(row, col).SharedFormulaID = int.MinValue;
                        }
                    }
                    removed.Add(index);
                }
            }
            foreach (int index in removed)
            {
                _worksheet._sharedFormulas.Remove(index);
            }
        }
        internal void SetSharedFormulaID(int id)
        {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _worksheet.Cell(row, col).SharedFormulaID = id;
                    }
                }
        }
        #endregion
        #region "Public Methods"
        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">print column names on first row</param>
        public void LoadFromDataTable(DataTable Table, bool PrintHeaders)
        {
            if (Table == null)
            {
                throw (new Exception("Table can't be null"));
            }

            int col = _fromCol, row = _fromRow;
            if (PrintHeaders)
            {
                foreach (DataColumn dc in Table.Columns)
                {
                    _worksheet.Cell(row, col++).Value = dc.ColumnName;
                }
                row++;
                col = _fromCol;
            }
            foreach (DataRow dr in Table.Rows)
            {
                foreach (object value in dr.ItemArray)
                {
                    _worksheet.Cell(row, col++).Value = value;
                }
                row++;
                col = _fromCol;
            }
        }
        /// <summary>
        /// Get a range with an offset from the top left cell.
        /// The new range has the same dimensions as the current range
        /// </summary>
        /// <param name="RowOffset">Row Offset</param>
        /// <param name="ColumnOffset">Column Offset</param>
        /// <returns></returns>
        public ExcelRangeBase Offset(int RowOffset, int ColumnOffset)
        {            
            if(_fromRow+RowOffset<1 || _fromCol+ColumnOffset<1 || _fromRow+RowOffset>ExcelPackage.MaxRows || _fromCol+ColumnOffset>ExcelPackage.MaxColumns)
            {
                throw(new ArgumentOutOfRangeException("Offset value out of range"));
            }
            string address = GetAddress(_fromRow+RowOffset, _fromCol+ColumnOffset, _toRow+RowOffset, _toCol+ColumnOffset);
            return new ExcelRangeBase(_worksheet, address);
        }
        /// <summary>
        /// Get a range with an offset from the top left cell.
        /// </summary>
        /// <param name="RowOffset">Row Offset</param>
        /// <param name="ColumnOffset">Column Offset</param>
        /// <param name="NumberOfRows">Number of rows. Minimum 1</param>
        /// <param name="NumberOfColumns">Number of colums. Minimum 1</param>
        /// <returns></returns>
        public ExcelRangeBase Offset(int RowOffset, int ColumnOffset, int NumberOfRows, int NumberOfColumns)
        {
            if (NumberOfRows < 1 || NumberOfColumns < 1)
            {
                throw (new Exception("Number of rows/columns must be greater than 0"));
            }
            NumberOfRows--;
            NumberOfColumns--;
            if(_fromRow+RowOffset<1 || _fromCol+ColumnOffset<1 || _fromRow+RowOffset>ExcelPackage.MaxRows || _fromCol+ColumnOffset>ExcelPackage.MaxColumns ||
               _fromRow+RowOffset+NumberOfRows<1 || _fromCol+ColumnOffset+NumberOfColumns<1 || _fromRow+RowOffset+NumberOfRows>ExcelPackage.MaxRows || _fromCol+ColumnOffset+NumberOfColumns>ExcelPackage.MaxColumns )
            {
                throw(new ArgumentOutOfRangeException("Offset value out of range"));
            }
            string address = GetAddress(_fromRow+RowOffset, _fromCol+ColumnOffset, _fromRow+RowOffset+NumberOfRows, _fromCol+ColumnOffset+NumberOfColumns);
            return new ExcelRangeBase(_worksheet, address);
        }
        /// <summary>
        /// Adds a new comment for the range.
        /// If this range contains more than one cell, the top left cell is returned by the method.
        /// </summary>
        /// <param name="Text"></param>
        /// <param name="Author"></param>
        /// <returns></returns>
        public ExcelComment AddComment(string Text, string Author)
        {
            ExistsComment();
            for (int col = _fromCol; col <= _toCol; col++)
            {
                for (int row = _fromRow; row <= _toRow; row++)
                {
                    Worksheet.Comments.Add(new ExcelRangeBase(_worksheet, GetAddress(_fromRow, _fromCol)), Text, Author);
                }
            }
            return  _worksheet.Cell(_fromRow, _fromCol).Comment;
        }

        private bool ExistsComment()
        {
            for (int col = _fromCol; col <= _toCol; col++)
            {
                for (int row = _fromRow; row <= _toRow; row++)
                {
                    if (_worksheet.Cell(row, col).Comment != null)
                    {
                        throw (new InvalidOperationException(string.Format("Cell {0} already contain a comment.", new ExcelCellAddress(row, col).Address)));
                    }
                }
            }
            return true;
        }
        #endregion
        #region IDisposable Members

        public void Dispose()
        {
            _worksheet = null;
        }

        #endregion
    }
}
