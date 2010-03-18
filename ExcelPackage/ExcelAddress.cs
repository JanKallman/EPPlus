using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// A range address
    /// </summary>
    public class ExcelAddress : ExcelCellBase
    {
        protected int _fromRow, _toRow, _fromCol, _toCol;
        protected string _address;
        #region "Constructors"
        internal ExcelAddress()
        {
        }
        public ExcelAddress(int fromRow, int fromCol, int toRow, int toColumn)
        {
            _fromRow = fromRow;
            _toRow = toRow;
            _fromCol = fromCol;
            _toCol = toColumn;
            _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
        }
        public ExcelAddress(string address)
        {
            Address = address;
        }
        ExcelCellAddress _start = null;
        #endregion
        /// <summary>
        /// Gets the row and column of the top left cell.
        /// </summary>
        /// <value>The start row column.</value>
        public ExcelCellAddress Start
        {
            get
            {
                if (_start == null)
                {
                    _start = new ExcelCellAddress(_fromRow, _fromCol);
                }
                return _end;
            }
        }
        ExcelCellAddress _end = null;
        /// <summary>
        /// Gets the row and column of the top left cell.
        /// </summary>
        /// <value>The start row column.</value>
        public ExcelCellAddress End
        {
            get
            {
                if (_end == null)
                {
                    _end = new ExcelCellAddress(_toRow, _fromRow);
                }
                return _end;
            }
        }
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
                GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
            }
        }
    }
}
