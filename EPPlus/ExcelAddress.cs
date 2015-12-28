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
 * Jan Källman		Added		18-MAR-2010
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml
{
    public class ExcelTableAddress
    {
        public string Name { get; set; }
        public string ColumnSpan { get; set; }
        public bool IsAll { get; set; }
        public bool IsHeader { get; set; }
        public bool IsData { get; set; }
        public bool IsTotals { get; set; }
        public bool IsThisRow { get; set; }
    }
    /// <summary>
    /// A range address
    /// </summary>
    /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
    public class ExcelAddressBase : ExcelCellBase
    {
        internal protected int _fromRow=-1, _toRow, _fromCol, _toCol;
        protected internal bool _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed;
        internal protected string _wb;
        internal protected string _ws;
        internal protected string _address;
        internal protected event EventHandler AddressChange;

        internal enum eAddressCollition
        {
            No,
            Partly,
            Inside,
            Equal
        }        
        internal enum eShiftType
        {
            Right,
            Down,
            EntireRow,
            EntireColumn
        }
        #region "Constructors"
        internal ExcelAddressBase()
        {
        }
        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <param name="fromRow">start row</param>
        /// <param name="fromCol">start column</param>
        /// <param name="toRow">End row</param>
        /// <param name="toColumn">End column</param>
        public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn)
        {
            _fromRow = fromRow;
            _toRow = toRow;
            _fromCol = fromCol;
            _toCol = toColumn;
            Validate();

            _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
        }
        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <param name="fromRow">start row</param>
        /// <param name="fromCol">start column</param>
        /// <param name="toRow">End row</param>
        /// <param name="toColumn">End column</param>
        /// <param name="fromRowFixed">start row fixed</param>
        /// <param name="fromColFixed">start column fixed</param>
        /// <param name="toRowFixed">End row fixed</param>
        /// <param name="toColFixed">End column fixed</param>
        public ExcelAddressBase(int fromRow, int fromCol, int toRow, int toColumn, bool fromRowFixed, bool fromColFixed, bool toRowFixed, bool toColFixed)
        {
            _fromRow = fromRow;
            _toRow = toRow;
            _fromCol = fromCol;
            _toCol = toColumn;
            _fromRowFixed = fromRowFixed;
            _fromColFixed = fromColFixed;
            _toRowFixed = toRowFixed;
            _toColFixed = toColFixed;
            Validate();

            _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, fromColFixed, _toRowFixed, _toColFixed );
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
        /// <summary>
        /// Creates an Address object
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        /// <param name="address">The Excel Address</param>
        /// <param name="pck">Reference to the package to find information about tables and names</param>
        /// <param name="referenceAddress">The address</param>
        public ExcelAddressBase(string address, ExcelPackage pck, ExcelAddressBase referenceAddress)
        {
            SetAddress(address);
            SetRCFromTable(pck, referenceAddress);
        }

        internal void SetRCFromTable(ExcelPackage pck, ExcelAddressBase referenceAddress)
        {
            if (string.IsNullOrEmpty(_wb) && Table != null)
            {
                foreach (var ws in pck.Workbook.Worksheets)
                {
                    foreach (var t in ws.Tables)
                    {
                        if (t.Name.Equals(Table.Name, StringComparison.InvariantCultureIgnoreCase))
                        {
                            _ws = ws.Name;
                            if (Table.IsAll)
                            {
                                _fromRow = t.Address._fromRow;
                                _toRow = t.Address._toRow;
                            }
                            else
                            {
                                if (Table.IsThisRow)
                                {
                                    if (referenceAddress == null)
                                    {
                                        _fromRow = -1;
                                        _toRow = -1;
                                    }
                                    else
                                    {
                                        _fromRow = referenceAddress._fromRow;
                                        _toRow = _fromRow;
                                    }
                                }
                                else if (Table.IsHeader && Table.IsData)
                                {
                                    _fromRow = t.Address._fromRow;
                                    _toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
                                }
                                else if (Table.IsData && Table.IsTotals)
                                {
                                    _fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
                                    _toRow = t.Address._toRow;
                                }
                                else if (Table.IsHeader)
                                {
                                    _fromRow = t.ShowHeader ? t.Address._fromRow : -1;
                                    _toRow = t.ShowHeader ? t.Address._fromRow : -1;
                                }
                                else if (Table.IsTotals)
                                {
                                    _fromRow = t.ShowTotal ? t.Address._toRow : -1;
                                    _toRow = t.ShowTotal ? t.Address._toRow : -1;
                                }
                                else
                                {
                                    _fromRow = t.ShowHeader ? t.Address._fromRow + 1 : t.Address._fromRow;
                                    _toRow = t.ShowTotal ? t.Address._toRow - 1 : t.Address._toRow;
                                }
                            }

                            if (string.IsNullOrEmpty(Table.ColumnSpan))
                            {
                                _fromCol = t.Address._fromCol;
                                _toCol = t.Address._toCol;
                                return;
                            }
                            else
                            {
                                var col = t.Address._fromCol;
                                var cols = Table.ColumnSpan.Split(':');
                                foreach (var c in t.Columns)
                                {
                                    if (_fromCol <= 0 && cols[0].Equals(c.Name, StringComparison.InvariantCultureIgnoreCase))   //Issue15063 Add invariant igore case
                                    {
                                        _fromCol = col;
                                        if (cols.Length == 1)
                                        {
                                            _toCol = _fromCol;
                                            return;
                                        }
                                    }
                                    else if (cols.Length > 1 && _fromCol > 0 && cols[1].Equals(c.Name, StringComparison.InvariantCultureIgnoreCase)) //Issue15063 Add invariant igore case
                                    {
                                        _toCol = col;
                                        return;
                                    }

                                    col++;
                                }
                            }
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// Address is an defined name
        /// </summary>
        /// <param name="address">the name</param>
        /// <param name="isName">Should always be true</param>
        internal ExcelAddressBase(string address, bool isName)
        {
            if (isName)
            {
                _address = address;
                _fromRow = -1;
                _fromCol = -1;
                _toRow = -1;
                _toCol = -1;
                _start = null;
                _end = null;
            }
            else
            {
                SetAddress(address);
            }
        }

        protected internal void SetAddress(string address)
        {
            address = address.Trim();
            if (address.StartsWith("'"))
            {
                int pos = address.IndexOf("'", 1);
                while (pos < address.Length && address[pos + 1] == '\'')
                {
                    pos = address.IndexOf("'", pos+2);
                }
                var wbws = address.Substring(1,pos-1).Replace("''","'");
                SetWbWs(wbws);
                _address = address.Substring(pos + 2);
            }
            else if (address.StartsWith("[")) //Remove any external reference
            {
                SetWbWs(address);
            }
            else
            {
                _address = address;
            }
            if(_address.IndexOfAny(new char[] {',','!', '['}) > -1)
            {
                //Advanced address. Including Sheet or multi or table.
                ExtractAddress(_address);
            }
            else
            {
                //Simple address
                GetRowColFromAddress(_address, out _fromRow, out _fromCol, out _toRow, out  _toCol, out _fromRowFixed, out _fromColFixed,  out _toRowFixed, out _toColFixed);
                _addresses = null;
                _start = null;
                _end = null;
            }
            _address = address;
            Validate();
        }
        internal void ChangeAddress()
        {
            if (AddressChange != null)
            {
                AddressChange(this, new EventArgs());
            }
        }
        private void SetWbWs(string address)
        {
            int pos;
            if (address[0] == '[')
            {
                pos = address.IndexOf("]");
                _wb = address.Substring(1, pos - 1);                
                _ws = address.Substring(pos + 1);
            }
            else
            {
                _wb = "";
                _ws = address;
            }
            pos = _ws.IndexOf("!");
            if (pos > -1)
            {
                _address = _ws.Substring(pos + 1);
                _ws = _ws.Substring(0, pos);
            }
        }
        internal void ChangeWorksheet(string wsName, string newWs)
        {
            if (_ws == wsName) _ws = newWs;
            var fullAddress = GetAddress();
            
            if (Addresses != null)
            {
                foreach (var a in Addresses)
                {
                    if (a._ws == wsName)
                    {
                        a._ws = newWs;
                        fullAddress += "," + a.GetAddress();
                    }
                    else
                    {
                        fullAddress += "," + a._address;
                    }
                }
            }
            _address = fullAddress;
        }

        private string GetAddress()
        {
            var adr = "";
            if (string.IsNullOrEmpty(_wb))
            {
                adr = "[" + _wb + "]";
            }

            if (string.IsNullOrEmpty(_ws))
            {
                adr += string.Format("'{0}'!", _ws);
            }
            adr += GetAddress(_fromRow, _fromCol, _toRow, _toCol);
            return adr;
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
        ExcelTableAddress _table=null;
        public ExcelTableAddress Table
        {
            get
            {
                return _table;
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
        /// If the address is a defined name
        /// </summary>
        public bool IsName
        {
            get
            {
                return _fromRow < 0;
            }
        }
        /// <summary>
        /// Returns the address text
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return _address;
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
        internal string AddressSpaceSeparated
        {
            get
            {
                return _address.Replace(',', ' '); //Conditional formatting and a few other places use space as separator for mulit addresses.
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
        internal protected List<ExcelAddress> _addresses = null;
        internal virtual List<ExcelAddress> Addresses
        {
            get
            {
                return _addresses;
            }
        }

        private bool ExtractAddress(string fullAddress)
        {
            var brackPos=new Stack<int>();
            var bracketParts=new List<string>();
            string first="", second="";
            bool isText=false, hasSheet=false;
            try
            {
                if (fullAddress == "#REF!")
                {
                    SetAddress(ref fullAddress, ref second, ref hasSheet);
                    return true;
                }
                else if (fullAddress.StartsWith("!"))
                {
                    // invalid address!
                    return false;
                }
                for (int i = 0; i < fullAddress.Length; i++)
                {
                    var c = fullAddress[i];
                    if (c == '\'')
                    {
                        if (isText && i + 1 < fullAddress.Length && fullAddress[i] == '\'')
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
                        isText = !isText;
                    }
                    else
                    {
                        if (brackPos.Count > 0)
                        {
                            if (c == '[' && !isText)
                            {
                                brackPos.Push(i);
                            }
                            else if (c == ']' && !isText)
                            {
                                if (brackPos.Count > 0)
                                {
                                    var from = brackPos.Pop();
                                    bracketParts.Add(fullAddress.Substring(from + 1, i - from - 1));

                                    if (brackPos.Count == 0)
                                    {
                                        HandleBrackets(first, second, bracketParts);
                                    }
                                }
                                else
                                {
                                    //Invalid address!
                                    return false;
                                }
                            }
                        }
                        else if (c == '[' && !isText)
                        {
                            brackPos.Push(i);
                        }
                        else if (c == '!' && !isText && !first.EndsWith("#REF") && !second.EndsWith("#REF"))
                        {
                            hasSheet = true;
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
                if (Table == null)
                {
                    SetAddress(ref first, ref second, ref hasSheet);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void HandleBrackets(string first, string second, List<string> bracketParts)
        {
            if(!string.IsNullOrEmpty(first))
            {
                _table = new ExcelTableAddress();
                Table.Name = first;
                foreach (var s in bracketParts)
                {
                    if(s.IndexOf("[")<0)
                    {
                        switch(s.ToLower(CultureInfo.InvariantCulture))                
                        {
                            case "#all":
                                _table.IsAll = true;
                                break;
                            case "#headers":
                               _table.IsHeader = true;
                                break;
                            case "#data":
                                _table.IsData = true;
                                break;
                            case "#totals":
                                _table.IsTotals = true;
                                break;
                            case "#this row":
                                _table.IsThisRow = true;
                                break;
                            default:
                                if(string.IsNullOrEmpty(_table.ColumnSpan))
                                {
                                    _table.ColumnSpan=s;
                                }
                                else
                                {
                                    _table.ColumnSpan += ":" + s;
                                }
                                break;
                        }                
                    }
                }
            }
        }
        #region Address manipulation methods
        internal eAddressCollition Collide(ExcelAddressBase address)
        {
            if (address.WorkSheet != WorkSheet && address.WorkSheet!=null)
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
        internal ExcelAddressBase AddRow(int row, int rows, bool setFixed=false)
        {
            if (row > _toRow)
            {
                return this;
            }
            else if (row <= _fromRow)
            {
                return new ExcelAddressBase((setFixed && _fromRowFixed ? _fromRow : _fromRow + rows), _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow + rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
            }
            else
            {
                return new ExcelAddressBase(_fromRow, _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow + rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
            }
        }
        internal ExcelAddressBase DeleteRow(int row, int rows, bool setFixed = false)
        {
            if (row > _toRow) //After
            {
                return this;
            }            
            else if (row+rows <= _fromRow) //Before
            {
                return new ExcelAddressBase((setFixed && _fromRowFixed ? _fromRow : _fromRow - rows), _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
            }
            else if (row <= _fromRow && row + rows > _toRow) //Inside
            {
                return null;
            }
            else  //Partly
            {
                if (row <= _fromRow)
                {
                    return new ExcelAddressBase(row, _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
                }
                else
                {
                    return new ExcelAddressBase(_fromRow, _fromCol, (setFixed && _toRowFixed ? _toRow : _toRow - rows < row ? row - 1 : _toRow - rows), _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
                }
            }
        }
        internal ExcelAddressBase AddColumn(int col, int cols, bool setFixed = false)
        {
            if (col > _toCol)
            {
                return this;
            }
            else if (col <= _fromCol)
            {
                return new ExcelAddressBase(_fromRow, (setFixed && _fromColFixed ? _fromCol : _fromCol + cols), _toRow, (setFixed && _toColFixed ? _toCol : _toCol + cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
            }
            else
            {
                return new ExcelAddressBase(_fromRow, _fromCol, _toRow, (setFixed && _toColFixed ? _toCol : _toCol + cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
            }
        }
        internal ExcelAddressBase DeleteColumn(int col, int cols, bool setFixed = false)
        {
            if (col > _toCol) //After
            {
                return this;
            }
            else if (col + cols <= _fromCol) //Before
            {
                return new ExcelAddressBase(_fromRow, (setFixed && _fromColFixed ? _fromCol : _fromCol - cols), _toRow, (setFixed && _toColFixed ? _toCol :_toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
            }
            else if (col <= _fromCol && col + cols > _toCol) //Inside
            {
                return null;
            }
            else  //Partly
            {
                if (col <= _fromCol)
                {
                    return new ExcelAddressBase(_fromRow, col, _toRow, (setFixed && _toColFixed ? _toCol : _toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
                }
                else
                {
                    return new ExcelAddressBase(_fromRow, _fromCol, _toRow, (setFixed && _toColFixed ? _toCol :_toCol - cols < col ? col - 1 : _toCol - cols), _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
                }
            }
        }
        internal ExcelAddressBase Insert(ExcelAddressBase address, eShiftType Shift/*, out ExcelAddressBase topAddress, out ExcelAddressBase leftAddress, out ExcelAddressBase rightAddress, out ExcelAddressBase bottomAddress*/)
        {
            //Before or after, no change
            //if ((_toRow > address._fromRow && _toCol > address.column) || 
            //    (_fromRow > address._toRow && column > address._toCol))
            if(_toRow < address._fromRow || _toCol < address._fromCol || (_fromRow > address._toRow && _fromCol > address._toCol))
            {
                //topAddress = null;
                //leftAddress = null;
                //rightAddress = null;
                //bottomAddress = null;
                return this;
            }

            int rows = address.Rows;
            int cols = address.Columns;
            string retAddress = "";
            if (Shift==eShiftType.Right)
            {
                if (address._fromRow > _fromRow)
                {
                    retAddress = GetAddress(_fromRow, _fromCol, address._fromRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
                }
                if(address._fromCol > _fromCol)
                {
                    retAddress = GetAddress(_fromRow < address._fromRow ? _fromRow : address._fromRow, _fromCol, address._fromRow, _toCol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
                }
            }
            if (_toRow < address._fromRow)
            {
                if (_fromRow < address._fromRow)
                {

                }
                else
                {
                }
            }
            return null;
        }
        #endregion
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
                if(string.IsNullOrEmpty(_ws) || !string.IsNullOrEmpty(ws)) _ws = ws;
                _firstAddress = address;
                GetRowColFromAddress(address, out _fromRow, out _fromCol, out _toRow, out  _toCol, out _fromRowFixed, out _fromColFixed, out _toRowFixed, out _toColFixed);
            }
            else
            {
                if (_addresses == null) _addresses = new List<ExcelAddress>();
                _addresses.Add(new ExcelAddress(_ws, address));
            }
        }
        internal enum AddressType
        {
            Invalid,
            InternalAddress,
            ExternalAddress,
            InternalName,
            ExternalName,
            Formula
        }

        internal static AddressType IsValid(string Address)
        {
            double d;
            if (Address == "#REF!")
            {
                return AddressType.Invalid;
            }
            else if(double.TryParse(Address, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) //A double, no valid address
            {
                return AddressType.Invalid;
            }
            else if (IsFormula(Address))
            {
                return AddressType.Formula;
            }
            else
            {
                string wb, ws, intAddress;
                if(SplitAddress(Address, out wb, out ws, out intAddress))
                {
                    if(intAddress.Contains("[")) //Table reference
                    {
                        return string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;
                    }
                    else if(intAddress.Contains(","))
                    {
                        intAddress=intAddress.Substring(0, intAddress.IndexOf(','));
                    }
                    if(IsAddress(intAddress))
                    {
                        return string.IsNullOrEmpty(wb) ? AddressType.InternalAddress : AddressType.ExternalAddress;
                    }
                    else
                    {
                        return string.IsNullOrEmpty(wb) ? AddressType.InternalName : AddressType.ExternalName;
                    }
                }
                else
                {
                    return AddressType.Invalid;
                }

                //if(string.IsNullOrEmpty(wb));

            }
            //ExcelAddress a = new ExcelAddress(Address);
            //if (Address.IndexOf('!') > 0)
            //{                
            //    string[] split = Address.Split('!');
            //    if (split.Length == 2)
            //    {
            //        ws = split[0];
            //        Address = split[1];
            //    }
            //    else if (split.Length == 3 && split[1] == "#REF" && split[2] == "")
            //    {
            //        ws = split[0];
            //        Address = "#REF!";
            //        if (ws.StartsWith("[") && ws.IndexOf("]") > 1)
            //        {
            //            return AddressType.ExternalAddress;
            //        }
            //        else
            //        {
            //            return AddressType.InternalAddress;
            //        }
            //    }
            //    else
            //    {
            //        return AddressType.Invalid;
            //    }            
            //}
            //int _fromRow, column, _toRow, _toCol;
            //if (ExcelAddressBase.GetRowColFromAddress(Address, out _fromRow, out column, out _toRow, out _toCol))
            //{
            //    if (_fromRow > 0 && column > 0 && _toRow <= ExcelPackage.MaxRows && _toCol <= ExcelPackage.MaxColumns)
            //    {
            //        if (ws.StartsWith("[") && ws.IndexOf("]") > 1)
            //        {
            //            return AddressType.ExternalAddress;
            //        }
            //        else
            //        {
            //            return AddressType.InternalAddress;
            //        }
            //    }
            //    else
            //    {
            //        return AddressType.Invalid;
            //    }
            //}
            //else
            //{
            //    if(IsValidName(Address))
            //    {
            //        if (ws.StartsWith("[") && ws.IndexOf("]") > 1)
            //        {
            //            return AddressType.ExternalName;
            //        }
            //        else
            //        {
            //            return AddressType.InternalName;
            //        }
            //    }
            //    else
            //    {
            //        return AddressType.Invalid;
            //    }
            //}

        }

        private static bool IsAddress(string intAddress)
        {
            if(string.IsNullOrEmpty(intAddress)) return false;            
            var cells = intAddress.Split(':');
            int fromRow,toRow, fromCol, toCol;

            if(!GetRowCol(cells[0], out fromRow, out fromCol, false))
            {
                return false;
            }
            if (cells.Length > 1)
            {
                if (!GetRowCol(cells[1], out toRow, out toCol, false))
                {
                    return false;
                }
            }
            else
            {
                toRow = fromRow;
                toCol = fromCol;
            }
            if( fromRow <= toRow && 
                fromCol <= toCol && 
                fromCol > -1 && 
                toCol <= ExcelPackage.MaxColumns && 
                fromRow > -1 && 
                toRow <= ExcelPackage.MaxRows)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private static bool SplitAddress(string Address, out string wb, out string ws, out string intAddress)
        {
            wb = "";
            ws = "";
            intAddress = "";
            var text = "";
            bool isText = false;
            var brackPos=-1;
            for (int i = 0; i < Address.Length; i++)
            {
                if (Address[i] == '\'')
                {
                    isText = !isText;
                    if(i>0 && Address[i-1]=='\'')
                    {
                        text += "'";
                    }
                }
                else
                {
                    if(Address[i]=='!' && !isText)
                    {
                        if (text.Length>0 && text[0] == '[')
                        {
                            wb = text.Substring(1, text.IndexOf("]") - 1);
                            ws = text.Substring(text.IndexOf("]") + 1);
                        }
                        else
                        {
                            ws=text;
                        }
                        intAddress=Address.Substring(i+1);
                        return true;
                    }
                    else
                    {
                        if(Address[i]=='[' && !isText)
                        {
                            if (i > 0) //Table reference return full address;
                            {
                                intAddress=Address;
                                return true;
                            }
                            brackPos=i;
                        }
                        else if(Address[i]==']' && !isText)
                        {
                            if (brackPos > -1)
                            {
                                wb = text;
                                text = "";
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            text+=Address[i];
                        }
                    }
                }
            }
            intAddress = text;
            return true;
        }

        private static bool IsFormula(string address)
        {
            var isText = false;
            for (int i = 0; i < address.Length; i++)
            {
                if (address[i] == '\'')
                {
                    isText = !isText;
                }
                else
                {
                    if (isText==false  && address.Substring(i, 1).IndexOfAny(new char[] { '(', ')', '+', '-', '*', '/', '=', '^', '&', '%', '\"' }) > -1)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private static bool IsValidName(string address)
        {
            if (Regex.IsMatch(address, "[^0-9./*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|][^/*-+,½!\"@#£%&/{}()\\[\\]=?`^~':;<>|]*"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public int Rows 
        {
            get
            {
                return _toRow - _fromRow+1;
            }
        }
        public int Columns
        {
            get
            {
                return _toCol - _fromCol + 1;
            }
        }

        internal bool IsMultiCell()
        {
            return (_fromRow < _fromCol || _fromCol < _toCol);
        }
        internal static String GetWorkbookPart(string address)
        {
            var ix = 0;
            if (address[0] == '[')
            {
                ix = address.IndexOf(']') + 1;
                if (ix > 0)
                {
                    return address.Substring(1, ix - 2);
                }
            }
            return "";
        }
        internal static string GetWorksheetPart(string address, string defaultWorkSheet)
        {
            int ix=0;
            return GetWorksheetPart(address, defaultWorkSheet, ref ix);
        }
        internal static string GetWorksheetPart(string address, string defaultWorkSheet, ref int endIx)
        {
            if(address=="") return defaultWorkSheet;
            var ix = 0;
            if (address[0] == '[')
            {
                ix = address.IndexOf(']')+1;
            }
            if (ix > 0 && ix < address.Length)
            {
                if (address[ix] == '\'')
                {
                    return GetString(address, ix, out endIx);
                }
                else
                {
                    var ixEnd = address.IndexOf('!',ix);
                    if(ixEnd>ix)
                    {
                        return address.Substring(ix, ixEnd-ix);
                    }
                    else
                    {
                        return defaultWorkSheet;
                    }
                }
            }
            else
            {
                return defaultWorkSheet;
            }
        }
        internal static string GetAddressPart(string address)
        {
            var ix=0;
            GetWorksheetPart(address, "", ref ix);
            if(ix<address.Length)
            {
                if (address[ix] == '!')
                {
                    return address.Substring(ix + 1);
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }

        }
        internal static void SplitAddress(string fullAddress, out string wb, out string ws, out string address, string defaultWorksheet="")
        {
            wb = GetWorkbookPart(fullAddress);
            int ix=0;
            ws = GetWorksheetPart(fullAddress, defaultWorksheet, ref ix);
            if (ix < fullAddress.Length)
            {
                if (fullAddress[ix] == '!')
                {
                    address = fullAddress.Substring(ix + 1);
                }
                else
                {
                    address = fullAddress.Substring(ix);
                }
            }
            else
            {
                address="";
            }
        }
        private static string GetString(string address, int ix, out int endIx)
        {
            var strIx = address.IndexOf("''");
            var prevStrIx = ix;
            while(strIx > -1) 
            {
                prevStrIx = strIx;
                strIx = address.IndexOf("''");
            }
            endIx = address.IndexOf("'");
            return address.Substring(ix, endIx - ix).Replace("''","'");
        }

        internal bool IsValidRowCol()
        {
            return !(_fromRow > _toRow  ||
                   _fromCol > _toCol ||
                   _fromRow < 1 ||
                   _fromCol < 1 ||
                   _toRow > ExcelPackage.MaxRows ||
                   _toCol > ExcelPackage.MaxColumns);
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
            if (string.IsNullOrEmpty(_ws)) _ws = ws;
        }
        internal ExcelAddress(string ws, string address, bool isName)
            : base(address, isName)
        {
            if (string.IsNullOrEmpty(_ws)) _ws = ws;
        }

        public ExcelAddress(string Address, ExcelPackage package, ExcelAddressBase referenceAddress) :
            base(Address, package, referenceAddress)
        {

        }
        /// <summary>
        /// The address for the range
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        public new string Address
        {
            get
            {
                if (string.IsNullOrEmpty(_address) && _fromRow>0)
                {
                    _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
                }
                return _address;
            }
            set
            {                
                SetAddress(value);
                base.ChangeAddress();
            }
        }
    }
    public class ExcelFormulaAddress : ExcelAddressBase
    {
        bool _fromRowFixed, _toRowFixed, _fromColFixed, _toColFixed;
        internal ExcelFormulaAddress()
            : base()
        {
        }

        public ExcelFormulaAddress(int fromRow, int fromCol, int toRow, int toColumn)
            : base(fromRow, fromCol, toRow, toColumn)
        {
            _ws = "";
        }
        public ExcelFormulaAddress(string address)
            : base(address)
        {
            SetFixed();
        }
        
        internal ExcelFormulaAddress(string ws, string address)
            : base(address)
        {
            if (string.IsNullOrEmpty(_ws)) _ws = ws;
            SetFixed();
        }
        internal ExcelFormulaAddress(string ws, string address, bool isName)
            : base(address, isName)
        {
            if (string.IsNullOrEmpty(_ws)) _ws = ws;
            if(!isName)
                SetFixed();
        }

        private void SetFixed()
        {
            if (Address.IndexOf("[") >= 0) return;
            var address=FirstAddress;
            if(_fromRow==_toRow && _fromCol==_toCol)
            {
                GetFixed(address, out _fromRowFixed, out _fromColFixed);
            }
            else
            {
                var cells = address.Split(':');
                GetFixed(cells[0], out _fromRowFixed, out _fromColFixed);
                GetFixed(cells[1], out _toRowFixed, out _toColFixed);
            }
        }

        private void GetFixed(string address, out bool rowFixed, out bool colFixed)
        {            
            rowFixed=colFixed=false;
            var ix=address.IndexOf('$');
            while(ix>-1)
            {
                ix++;
                if(ix < address.Length)
                {
                    if(address[ix]>='0' && address[ix]<='9')
                    {
                        rowFixed=true;
                        break;
                    }
                    else
                    {
                        colFixed=true;
                    }
                }
                ix = address.IndexOf('$', ix);
            }
        }
        /// <summary>
        /// The address for the range
        /// </summary>
        /// <remarks>Examples of addresses are "A1" "B1:C2" "A:A" "1:1" "A1:E2,G3:G5" </remarks>
        public new string Address
        {
            get
            {
                if (string.IsNullOrEmpty(_address) && _fromRow>0)
                {
                    _address = GetAddress(_fromRow, _fromCol, _toRow, _toCol, _fromRowFixed, _toRowFixed, _fromColFixed, _toColFixed);
                }
                return _address;
            }
            set
            {                
                SetAddress(value);
                base.ChangeAddress();
                SetFixed();
            }
        }
        internal new List<ExcelFormulaAddress> _addresses;
        public new List<ExcelFormulaAddress> Addresses
        {
            get
            {
                if (_addresses == null)
                {
                    _addresses = new List<ExcelFormulaAddress>();
                }
                return _addresses;

            }
        }
        internal string GetOffset(int row, int column)
        {
            int fromRow = _fromRow, fromCol = _fromCol, toRow = _toRow, tocol = _toCol;
            var isMulti = (fromRow != toRow || fromCol != tocol);
            if (!_fromRowFixed)
            {
                fromRow += row;
            }
            if (!_fromColFixed)
            {
                fromCol += column;
            }
            if (isMulti)
            {
                if (!_toRowFixed)
                {
                    toRow += row;
                }
                if (!_toColFixed)
                {
                    tocol += column;
                }
            }
            else
            {
                toRow = fromRow;
                tocol = fromCol;
            }
            string a = GetAddress(fromRow, fromCol, toRow, tocol, _fromRowFixed, _fromColFixed, _toRowFixed, _toColFixed);
            if (Addresses != null)
            {
                foreach (var sa in Addresses)
                {
                    a+="," + sa.GetOffset(row, column);
                }
            }
            return a;
        }
    }
}
