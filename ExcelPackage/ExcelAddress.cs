/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * See http://epplus.codeplex.com/ for details
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
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		18-MAR-2010
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// A range address
    /// </summary>
    /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
    public class ExcelAddressBase : ExcelCellBase
    {
        internal protected int _fromRow=-1, _toRow, _fromCol, _toCol;
        internal protected string _ws;
        internal protected string _address;
        internal enum eAddressCollition
        {
            No,
            Partly,
            Inside,
            Equal
        }
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
//            GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
        }
        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        /// <param name="address">The Excel Address</param>
        public ExcelAddressBase(string address)
        {
            SetAddress(address);
        }

        protected internal void SetAddress(string address)
        {
            _address = address;
            if(address.IndexOfAny(new char[] {',','!'}) > -1)
            {
                //Advanced address. Including Sheet or multi
                ExtractAddress(_address);
            }
            else
            {
                //Simple address
                GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
                _addresses = null;
                _start = null;
                _end = null;
            }
            Validate();
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
        string _firstAddress;
        /// <summary>
        /// returns the first address if the address is a multi address.
        /// A1:A2,B1:B2 returns A1:A2
        /// </summary>
        internal string FirstAddress
        {
            get
            {
                if (string.IsNullOrEmpty(_firstAddress))
                {
                    return _address;
                }
                else
                {
                    return _firstAddress;
                }
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
        internal string WorkSheet
        {
            get
            {
                return _ws;
            }
        }
        List<ExcelAddress> _addresses = null;
        internal List<ExcelAddress> Addresses
        {
            get
            {
                return _addresses;
            }
        }

        private void ExtractAddress(string fullAddress)
        {
            string first="", second="";
            bool isText=false, hasSheet=false;
            if (fullAddress == "#REF!")
            {
                SetAddress(ref fullAddress, ref second, ref hasSheet );
                return;
            }
            foreach (char c in fullAddress)
            {
                if(c=='\'')
                {
                    isText=!isText;
                }
                else
                {
                    if(c=='!' && !isText && !first.EndsWith("#REF") && !second.EndsWith("#REF"))
                    {
                        hasSheet=true;
                    }
                    else if (c == ',' && !isText)
                    {
                        SetAddress(ref first, ref second, ref hasSheet);
                    }
                    else
                    {
                        if (hasSheet)
                        {
                            second += c;
                        }
                        else
                        {
                            first += c;
                        }
                    }
                }
            }
            SetAddress(ref first, ref second, ref hasSheet);
        }
        internal eAddressCollition Collide(ExcelAddressBase address)
        {
            if (address.WorkSheet != WorkSheet)
            {
                return eAddressCollition.No;
            }

            if (address._fromRow > _toRow || address._fromCol > _toCol
                ||
                _fromRow > address._toRow || _fromCol > address._toCol)
            {
                return eAddressCollition.No;
            }
            else if (address._fromRow == _fromRow && address._fromCol == _fromCol &&
                    address._toRow == _toRow && address._toCol == _toCol)
            {
                return eAddressCollition.Equal;
            }
            else if (address._fromRow >= _fromRow && address._toRow <= _toRow &&
                     address._fromCol >= _fromCol && address._toCol <= _toCol)
            {
                return eAddressCollition.Inside;
            }
            else
                return eAddressCollition.Partly;
        }
        private void SetAddress(ref string first, ref string second, ref bool hasSheet)
        {
            string ws, address;
            if (hasSheet)
            {
                ws = first;
                address = second;
                first = "";
                second = "";
            }
            else
            {
                address = first;
                ws = "";
                first = "";
            }
            hasSheet = false;
            if (string.IsNullOrEmpty(_firstAddress))
            {
                _ws = ws;
                _firstAddress = address;
                GetRowColFromAddress(address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
            }
            else
            {
                if (_addresses == null) _addresses = new List<ExcelAddress>();
                _addresses.Add(new ExcelAddress(_ws, address));
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
            _ws = "";
        }
        public ExcelAddress(string address)
            : base(address)
        {
        }
        
        internal ExcelAddress(string ws, string address)
            : base(address)
        {
            _ws = ws;
        }
        /// <summary>
        /// The address for the range
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        public string Address
        {
            get
            {
                return _address;
            }
            set
            {
                SetAddress(value);
                //base.Address = value;
                //_address = value;
                //GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol);
                //Validate();
            }
        }
    }
}
