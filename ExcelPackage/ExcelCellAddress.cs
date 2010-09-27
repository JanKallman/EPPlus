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
 *  Starnuto Di Topo & Jan Källman  Initial Release		        2010-03-14
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// A single Cell address 
    /// </summary>
    public class ExcelCellAddress
    {
        public ExcelCellAddress()
            : this(1, 1)
        {

        }

        private int _row;
        private int _column;
        private string _address;
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelCellPosition"/> class.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        public ExcelCellAddress(int row, int column)
        {
            this.Row = row;
            this.Column = column;
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelCellPosition"/> class.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        public ExcelCellAddress(string address)
        {
            this.Address = address; 
        }
        /// <summary>
        /// Row
        /// </summary>
        public int Row
        {
            get
            {
                return this._row;
            }
            private set
            {
                if (value <= 0)
                {
                    throw new ArgumentOutOfRangeException("value", "Row cannot be less then 1.");
                }
                this._row = value;
                if(_column>0) 
                    _address = ExcelCellBase.GetAddress(_row, _column);
                else
                    _address = "#REF!";
            }
        }
        /// <summary>
        /// Column
        /// </summary>
        public int Column
        {
            get
            {
                return this._column;
            }
            private set
            {
                if (value <= 0)
                {
                    throw new ArgumentOutOfRangeException("value", "Column cannot be less then 1.");
                }
                this._column = value;
                if (_row > 0)
                    _address = ExcelCellBase.GetAddress(_row, _column);
                else
                    _address = "#REF!";
            }
        }
        /// <summary>
        /// Celladdress
        /// </summary>
        public string Address
        {
            get
            {
                return _address;
            }
            private set
            {
                _address = value;
                ExcelCellBase.GetRowColFromAddress(_address, out _row, out _column);
            }
        }
        public bool IsRef
        {
            get
            {
                return _row <= 0;
            }
        }
    }
}

