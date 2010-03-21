using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// A range address
    /// </summary>
    public class ExcelAddressBase : ExcelCellBase
    {
        protected int _fromRow, _toRow, _fromCol, _toCol;
        protected string _address;
        #region "Constructors"
        internal ExcelAddressBase()
        {
        }
        public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn)
        {
            _fromRow = fromRow;
            _toRow = toRow;
            _fromCol = fromCol;
            _toCol = toColumn;
            Validate();

            _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
            GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
        }
        public ExcelAddressBase(string address)
        {
            _address = address;
            GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
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
                return _start;
            }
        }
        ExcelCellAddress _end = null;
        /// <summary>
        /// Gets the row and column of the bottom right cell.
        /// </summary>
        /// <value>The end row column.</value>
        public ExcelCellAddress End
        {
            get
            {
                if (_end == null)
                {
                    _end = new ExcelCellAddress(_toRow, _toCol);
                }
                return _end;
            }
        }
        /// <summary>
        /// The address for the range
        /// </summary>
        public virtual string Address
        {
            get
            {
                return _address;
            }
        }

        /// <summary>
        /// Validate the address
        /// </summary>
        protected void Validate()
        {
            if (_fromRow > _toRow || _fromCol > _toCol)
            {
                throw new ArgumentOutOfRangeException("Start cell Address must be less or equal to End cell address");
            }
        }
    }
    /// <summary>
    /// Range address with the address property readonly
    /// </summary>
    public class ExcelAddress : ExcelAddressBase
    {
        internal ExcelAddress()
            : base()
        {

        }

        public ExcelAddress(int fromRow, int fromCol, int toRow, int toColumn)
            : base(fromRow, fromCol, toRow, toColumn)
        {

        }
        public ExcelAddress(string address)
        {
            Validate();
            GetRowColFromAddress(address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
            Address = address;
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
                Validate();
            }
        }
    }
}
