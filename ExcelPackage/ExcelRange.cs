/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * EPPlus is a fork of the ExcelPackage project
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
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Style;
using System.Data;
namespace OfficeOpenXml
{
    public class ExcelRange : ExcelCellBase, IExcelCell, IDisposable
    {
        private ExcelWorksheet _xlWorksheet;
        private int _fromRow = 1, _toRow = 1, _fromCol = 1, _toCol = 1;
        string _address;

        #region "Constructors"
        protected internal ExcelRange(ExcelWorksheet xlWorksheet)
        {
            _xlWorksheet = xlWorksheet;
            if (_xlWorksheet.View.SelectedRange == "")
            {
                _address = "A1";
                return;
            }
            else
            {
                _address = _xlWorksheet.View.SelectedRange;
            }
            GetAddressRowCol(_address, out _fromCol, out  _fromRow, out  _toCol, out _toRow);
            //_address = Address;
            //GetRangeRowCol(_address, out _fromCol, out  _fromRow, out  _toCol, out _toRow);
        }
        #endregion
        #region "Indexers"
        /// <summary>
        /// Access the range using an address
        /// </summary>
        /// <param name="Address">The address</param>
        /// <returns>A range object</returns>
        public ExcelRange this[string Address]
        {
            get
            {
                _address = Address;
                GetAddressRowCol(_address, out _fromCol, out  _fromRow, out  _toCol, out _toRow);
                return this;
            }
        }
        /// <summary>
        /// Access a single cell
        /// </summary>
        /// <param name="Row">The row</param>
        /// <param name="Col">The column</param>
        /// <returns>A range object</returns>
        public ExcelRange this[int Row, int Col]
        {
            get
            {
                _fromCol = Col;
                _fromRow = Row;
                _toCol = Col;
                _toRow = Row;
                _address = GetAddress(_fromRow, _fromCol);
                return this;
            }
        }
        /// <summary>
        /// Access a range of cells
        /// </summary>
        /// <param name="FromRow">Start row</param>
        /// <param name="FromCol">Start column</param>
        /// <param name="ToRow">End Row</param>
        /// <param name="ToCol">End Column</param>
        /// <returns></returns>
        public ExcelRange this[int FromRow, int FromCol, int ToRow, int ToCol]
        {
            get
            {
                _fromCol = FromCol;
                _fromRow = FromRow;
                _toCol = ToCol;
                _toRow = ToRow;
                _address = GetAddress(_fromRow, _fromCol) + ":" + ExcelCell.GetAddress(_toRow, _toCol);
                return this;
            }
        }
        #endregion
        #region "Public Properties"
        /// <summary>
        /// The address for the range
        /// </summary>
        public string Address
        {
            get
            {
                return _address;
            }
            set
            {
                _address = value;
                GetAddressRowCol(_address, out _fromCol, out  _fromRow, out  _toCol, out _toRow);
            }
        }
        /// <summary>
        /// The styleobject for the range.
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return _xlWorksheet.Workbook.Styles.GetStyleObject(_xlWorksheet.Cell(_fromRow, _fromCol).StyleID ,_xlWorksheet.PositionID, _address);
            }
        }
        /// <summary>
        /// The named style
        /// </summary>
        public string StyleName
        {
            get
            {
                return _xlWorksheet.Cell(_fromRow, _fromCol).StyleName;
            }
            set
            {
                int styleID = _xlWorksheet.Workbook.Styles.GetStyleIdFromName(value);
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _xlWorksheet.Cell(row, col).SetNewStyleName(value, styleID);
                    }
                }
            }
        }
        /// <summary>
        /// The style ID. Can be used for fast copying of styles
        /// </summary>
        public int StyleID
        {
            get
            {
                return _xlWorksheet.Cell(_fromRow, _fromCol).StyleID;
            }
            set
            {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _xlWorksheet.Cell(_fromRow, _fromCol).StyleID = value;
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
                return _xlWorksheet.Cell(_fromRow, _fromCol).Value;
            }
            set
            {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _xlWorksheet.Cell(row, col).Value=value;
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
                return _xlWorksheet.Cell(_fromRow, _fromCol).Formula;
            }
            set
            {
                if (value[0] == '=') value = value.Substring(1, value.Length - 1); // remove any starting equalsign.
                RemoveFormuls();
                ExcelWorksheet.Formulas f = new ExcelWorksheet.Formulas();
                f.Formula = value;
                f.Index = _xlWorksheet.GetMaxShareFunctionIndex();
                f.Address = _address;
                f.StartCol = _fromCol;
                f.StartRow = _fromRow;

                _xlWorksheet._sharedFormulas.Add(f.Index, f);
                _xlWorksheet.Cell(_fromRow, _fromCol).SharedFormulaID = f.Index;
                _xlWorksheet.Cell(_fromRow, _fromCol).Formula = value;

                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _xlWorksheet.Cell(row, col).SharedFormulaID = f.Index;
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
                return _xlWorksheet.Cell(_fromRow, _fromCol).FormulaR1C1;
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
                return _xlWorksheet.Cell(_fromRow, _fromCol).Hyperlink;
            }
            set
            {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _xlWorksheet.Cell(row, col).Hyperlink = value;
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
                        if (!_xlWorksheet.Cell(row, col).Merge)
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
                    if (_xlWorksheet.MergedCells.List.Contains(_address))
                    {
                        SetCellMerge(false);
                        _xlWorksheet.MergedCells.List.Remove(_address);
                    }
                    else if(!CheckMergeDiff(false))
                    {
                        throw (new Exception("Range is not fully merged.Specify the exact range"));                        
                    }
                }
                else
                {
                    if (CheckMergeDiff(false))
                    {
                        SetCellMerge(true);
                        _xlWorksheet.MergedCells.List.Add(_address);
                    }
                    else
                    {
                        if (!_xlWorksheet.MergedCells.List.Contains(_address))
                        {
                            throw(new Exception("Cells are already merged"));
                        }
                    }
                }
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
                return _xlWorksheet.Cell(_fromRow, _fromCol).IsRichText;   
            }
            set
            {
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        _xlWorksheet.Cell(row, col).IsRichText = value;
                    }
                }
            }
        }
        #endregion
        #region "Private Functions"
        /// <summary>
        /// Check if the range is partly merged
        /// </summary>
        /// <returns></returns>
        private bool CheckMergeDiff()
        {
            return CheckMergeDiff(_xlWorksheet.Cell(_fromRow, _fromCol).Merge);
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
                    if (_xlWorksheet.Cell(row, col).Merge != startValue)
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
        private void SetCellMerge(bool value)
        {
            for (int col = _fromCol; col <= _toCol; col++)
            {
                for (int row = _fromRow; row <= _toRow; row++)
                {
                    _xlWorksheet.Cell(row, col).Merge = value;
                }
            }
        }
        /// <summary>
        /// Removes a shared formula
        /// </summary>
        private void RemoveFormuls()
        {
            List<int> removed=new List<int>();
            int fFromRow, fFromCol, fToRow, fToCol;
            foreach (int index in _xlWorksheet._sharedFormulas.Keys)
            {
                ExcelWorksheet.Formulas f = _xlWorksheet._sharedFormulas[index];
                ExcelCell.GetAddressRowCol(f.Address, out fFromCol, out fFromRow, out fToCol, out fToRow);
                if(((fFromCol >= _fromCol && fFromCol <= _toCol) ||
                   (fToCol >= _fromCol && fToCol <= _toCol)) && 
                   ((fFromRow >= _fromRow && fFromRow <= _toRow) ||
                   (fToRow >= _fromRow && fToRow <= _toRow)))
                {
                    for (int col = fFromCol; col <= fToCol; col++)
                    {
                        for (int row = fFromRow; row <= fToRow; row++)
                        {
                            _xlWorksheet.Cell(row,col).SharedFormulaID = int.MinValue;
                        }
                    }
                    removed.Add(index);
                }
            }
            foreach (int index in removed)
            {
                _xlWorksheet._sharedFormulas.Remove(index);
            }
        }
        #endregion
        #region "Public Functions"
        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">print column names on first row</param>
        public void LoadFromDataTable(DataTable Table, bool PrintHeaders)
        {
            if (Table == null)
            {
                throw(new Exception("Table can't be null"));
            }

            int col = _fromCol, row = _fromRow;
            if (PrintHeaders)
            {
                foreach (DataColumn dc in Table.Columns)
                {
                    _xlWorksheet.Cell(row, col++).Value = dc.ColumnName;
                }
                row++;
                col=_fromCol;
            }
            foreach (DataRow dr in Table.Rows)
            {                
                foreach(object value in dr.ItemArray)
                {
                    _xlWorksheet.Cell(row, col++).Value = value;
                }
                row++;
                col = _fromCol;
            }
        }
        #endregion
        #region IDisposable Members

        public void Dispose()
        {
            _xlWorksheet = null;
        }

        #endregion
    }
}
