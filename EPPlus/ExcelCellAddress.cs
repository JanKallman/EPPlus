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
 *******************************************************************************
 *  Starnuto Di Topo & Jan Källman  Initial Release		        2010-03-14
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// A single cell address 
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
        /// Initializes a new instance of the ExcelCellAddress class.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        public ExcelCellAddress(int row, int column)
        {
            this.Row = row;
            this.Column = column;
        }
        /// <summary>
        /// Initializes a new instance of the ExcelCellAddress class.
        /// </summary>
        ///<param name="address">The address</param>
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
                    throw new ArgumentOutOfRangeException("value", "Row cannot be less than 1.");
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
                    throw new ArgumentOutOfRangeException("value", "Column cannot be less than 1.");
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
            internal set
            {
                _address = value;
                ExcelCellBase.GetRowColFromAddress(_address, out _row, out _column);
            }
        }
        /// <summary>
        /// If the address is an invalid reference (#REF!)
        /// </summary>
        public bool IsRef
        {
            get
            {
                return _row <= 0;
            }
        }

        /// <summary>
        /// Returns the letter corresponding to the supplied 1-based column index.
        /// </summary>
        /// <param name="column">Index of the column (1-based)</param>
        /// <returns>The corresponding letter, like A for 1.</returns>
        public static string GetColumnLetter(int column)
        {
            if (column > ExcelPackage.MaxColumns || column < 1)
            {
                throw new InvalidOperationException("Invalid 1-based column index: " + column + ". Valid range is 1 to " + ExcelPackage.MaxColumns);
            }
            return ExcelCellBase.GetColumnLetter(column);
        }
    }
}

