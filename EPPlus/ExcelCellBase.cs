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
 * Jan Källman		Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing;
namespace OfficeOpenXml
{
    /// <summary>
    /// Base class containing cell address manipulating methods.
    /// </summary>
    public abstract class ExcelCellBase
    {
        #region "public functions"
        /// <summary>
        /// Get the sheet, row and column from the CellID
        /// </summary>
        /// <param name="cellID"></param>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        static internal void SplitCellID(ulong cellID, out int sheet, out int row, out int col)
        {
            sheet = (int)(cellID % 0x8000);
            col = ((int)(cellID >> 15) & 0x3FF);
            row = ((int)(cellID >> 29));
        }
        /// <summary>
        /// Get the cellID for the cell. 
        /// </summary>
        /// <param name="SheetID"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        internal static ulong GetCellID(int SheetID, int row, int col)
        {
            return ((ulong)SheetID) + (((ulong)col) << 15) + (((ulong)row) << 29);
        }
        #endregion
        #region "Formula Functions"
        private delegate string dlgTransl(string part, int row, int col, int rowIncr, int colIncr);
        #region R1C1 Functions"
        /// <summary>
        /// Translates a R1C1 to an absolut address/Formula
        /// </summary>
        /// <param name="value">Address</param>
        /// <param name="row">Current row</param>
        /// <param name="col">Current column</param>
        /// <returns>The RC address</returns>
        public static string TranslateFromR1C1(string value, int row, int col)
        {
            return Translate(value, ToAbs, row, col, -1, -1);
        }
        /// <summary>
        /// Translates a absolut address to R1C1 Format
        /// </summary>
        /// <param name="value">R1C1 Address</param>
        /// <param name="row">Current row</param>
        /// <param name="col">Current column</param>
        /// <returns>The absolut address/Formula</returns>
        public static string TranslateToR1C1(string value, int row, int col)
        {
            return Translate(value, ToR1C1, row, col, -1, -1);
        }
        /// <summary>
        /// Translates betweein R1C1 or absolut addresses
        /// </summary>
        /// <param name="value">The addresss/function</param>
        /// <param name="addressTranslator">The translating function</param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="rowIncr"></param>
        /// <param name="colIncr"></param>
        /// <returns></returns>
        private static string Translate(string value, dlgTransl addressTranslator, int row, int col, int rowIncr, int colIncr)
        {
            if (value == "")
                return "";
            bool isText = false;
            string ret = "";
            string part = "";
            char prevTQ = (char)0;
            for (int pos = 0; pos < value.Length; pos++)
            {
                char c = value[pos];
                if (c == '"' || c=='\'')
                {
                    if (isText == false && part != "" && prevTQ==c)
                    {
                        ret += addressTranslator(part, row, col, rowIncr, colIncr);
                        part = "";
                        prevTQ = (char)0;
                    }
                    prevTQ = c;
                    isText = !isText;
                    ret += c;
                }
                else if (isText)
                {
                    ret += c;
                }
                else
                {
                    if ((c == '-' || c == '+' || c == '*' || c == '/' ||
                        c == '=' || c == '^' || c == ',' || c == ':' ||
                        c == '<' || c == '>' || c == '(' || c == ')' || c == '!' ||
                        c == ' ' || c == '&' || c == '%') &&
                        (pos == 0 || value[pos - 1] != '[')) //Last part to allow for R1C1 style [-x]
                    {
                        ret += addressTranslator(part, row, col, rowIncr, colIncr) + c;
                        part = "";
                    }
                    else
                    {
                        part += c;
                    }
                }
            }
            if (part != "")
            {
                ret += addressTranslator(part, row, col, rowIncr, colIncr);
            }
            return ret;
        }
        /// <summary>
        /// Translate to R1C1
        /// </summary>
        /// <param name="part">the value to be translated</param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="rowIncr"></param>
        /// <param name="colIncr"></param>
        /// <returns></returns>
        private static string ToR1C1(string part, int row, int col, int rowIncr, int colIncr)
        {
            int addrRow, addrCol;
            string Ret = "R";
            if (GetRowCol(part, out addrRow, out addrCol, false))
            {
                if (addrRow == 0 || addrCol == 0)
                {
                    return part;
                }
                if (part.IndexOf('$', 1) > 0)
                {
                    Ret += addrRow.ToString();
                }
                else if (addrRow - row != 0)
                {
                    Ret += string.Format("[{0}]", addrRow - row);
                }

                if (part.StartsWith("$"))
                {
                    return Ret + "C" + addrCol;
                }
                else if (addrCol - col != 0)
                {
                    return Ret + "C" + string.Format("[{0}]", addrCol - col);
                }
                else
                {
                    return Ret + "C";
                }
            }
            else
            {
                return part;
            }
        }
        /// <summary>
        /// Translates to absolute address
        /// </summary>
        /// <param name="part"></param>
        /// <param name="row"></param>
        /// <param name="col"></param> 
        /// <param name="rowIncr"></param>
        /// <param name="colIncr"></param>
        /// <returns></returns>
        private static string ToAbs(string part, int row, int col, int rowIncr, int colIncr)
        {
            string check = part.ToUpper(CultureInfo.InvariantCulture);

            int rStart = check.IndexOf("R");
            if (rStart != 0)
                return part;
            if (part.Length == 1) //R
            {
                return GetAddress(row, col);
            }

            int cStart = check.IndexOf("C");
            bool absoluteRow, absoluteCol;
            if (cStart == -1)
            {
                int RNum = GetRC(part, row, out absoluteRow);
                if (RNum > int.MinValue)
                {
                    return GetAddress(RNum, absoluteRow, col, false);
                }
                else
                {
                    return part;
                }
            }
            else
            {
                int RNum = GetRC(part.Substring(1, cStart - 1), row, out absoluteRow);
                int CNum = GetRC(part.Substring(cStart + 1, part.Length - cStart - 1), col, out absoluteCol);
                if (RNum > int.MinValue && CNum > int.MinValue)
                {
                    return GetAddress(RNum, absoluteRow, CNum, absoluteCol);
                }
                else
                {
                    return part;
                }
            }
        }
        /// <summary>
        /// Adds or subtracts a row or column to an address
        /// </summary>
        /// <param name="Address"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="rowIncr"></param>
        /// <param name="colIncr"></param>
        /// <returns></returns>
        private static string AddToRowColumnTranslator(string Address, int row, int col, int rowIncr, int colIncr)
        {
            int fromRow, fromCol;
            if (Address == "#REF!")
            {
                return Address;
            }
            if (GetRowCol(Address, out fromRow, out fromCol, false))
            {
                if (fromRow == 0 || fromCol == 0)
                {
                    return Address;
                }
                if (rowIncr != 0 && row != 0 && fromRow >= row && Address.IndexOf('$', 1) == -1)
                {
                    if (fromRow < row - rowIncr)
                    {
                        return "#REF!";
                    }

                    fromRow = fromRow + rowIncr;
                }

                if (colIncr != 0 && col != 0 && fromCol >= col && Address.StartsWith("$") == false)
                {
                    if (fromCol < col - colIncr)
                    {
                        return "#REF!";
                    }

                    fromCol = fromCol + colIncr;
                }

                Address = GetAddress(fromRow, Address.IndexOf('$', 1) > -1, fromCol, Address.StartsWith("$"));
            }
            return Address;
        }

        /// <summary>
        /// Returns with brackets if the value is negative
        /// </summary>
        /// <param name="v">The value</param>
        /// <returns></returns>
        private static string GetRCFmt(int v)
        {
            return (v < 0 ? string.Format("[{0}]", v) : v > 0 ? v.ToString() : "");
        }
        /// <summary>
        /// Get the offset value for RC format
        /// </summary>
        /// <param name="value"></param>
        /// <param name="OffsetValue"></param>
        /// <param name="fixedAddr"></param>
        /// <returns></returns>
        private static int GetRC(string value, int OffsetValue, out bool fixedAddr)
        {
            if (value == "")
            {
                fixedAddr = false;
                return OffsetValue;
            }
            int num;
            if (value[0] == '[' && value[value.Length - 1] == ']') //Offset?                
            {
                fixedAddr = false;
                if (int.TryParse(value.Substring(1, value.Length - 2), out num))
                {
                    return (OffsetValue + num);
                }
                else
                {
                    return int.MinValue;
                }
            }
            else
            {
                fixedAddr = true;
                if (int.TryParse(value, out num))
                {
                    return num;
                }
                else
                {
                    return int.MinValue;
                }
            }
        }
        #endregion
        #region "Address Functions"
        #region GetColumnLetter
        /// <summary>
        /// Returns the character representation of the numbered column
        /// </summary>
        /// <param name="iColumnNumber">The number of the column</param>
        /// <returns>The letter representing the column</returns>
        protected internal static string GetColumnLetter(int iColumnNumber)
        {
            return GetColumnLetter(iColumnNumber, false);
        }
        protected internal static string GetColumnLetter(int iColumnNumber, bool fixedCol)
        {

            if (iColumnNumber < 1)
            {
                //throw new Exception("Column number is out of range");
                return "#REF!";
            }

            string sCol = "";
            do
            {
                sCol = ((char)('A' + ((iColumnNumber - 1) % 26))) + sCol;
                iColumnNumber = (iColumnNumber - ((iColumnNumber - 1) % 26)) / 26;
            }
            while (iColumnNumber > 0);
            return fixedCol ? "$" + sCol : sCol;
        }
        #endregion

        internal static bool GetRowColFromAddress(string CellAddress, out int FromRow, out int FromColumn, out int ToRow, out int ToColumn)
        {
            bool fixedFromRow, fixedFromColumn, fixedToRow, fixedToColumn;
            return GetRowColFromAddress(CellAddress, out FromRow, out FromColumn, out ToRow, out ToColumn, out fixedFromRow, out fixedFromColumn, out fixedToRow, out fixedToColumn);
        }
        /// <summary>
        /// Get the row/columns for a Cell-address
        /// </summary>
        /// <param name="CellAddress">The address</param>
        /// <param name="FromRow">Returns the to column</param>
        /// <param name="FromColumn">Returns the from column</param>
        /// <param name="ToRow">Returns the to row</param>
        /// <param name="ToColumn">Returns the from row</param>
        /// <param name="fixedFromRow">Is the from row fixed?</param>
        /// <param name="fixedFromColumn">Is the from column fixed?</param>
        /// <param name="fixedToRow">Is the to row fixed?</param>
        /// <param name="fixedToColumn">Is the to column fixed?</param>
        /// <returns></returns>
        internal static bool GetRowColFromAddress(string CellAddress, out int FromRow, out int FromColumn, out int ToRow, out int ToColumn, out bool fixedFromRow, out bool fixedFromColumn, out bool fixedToRow, out bool fixedToColumn)
        {
            bool ret;
            if (CellAddress.IndexOf('[') > 0) //External reference or reference to Table or Pivottable.
            {
                FromRow = -1;
                FromColumn = -1;
                ToRow = -1;
                ToColumn = -1;
                fixedFromRow = false;
                fixedFromColumn = false;
                fixedToRow= false;
                fixedToColumn = false;
                return false;
            }

            CellAddress = CellAddress.ToUpper(CultureInfo.InvariantCulture);
            //This one can be removed when the worksheet Select format is fixed
            if (CellAddress.IndexOf(' ') > 0)
            {
                CellAddress = CellAddress.Substring(0, CellAddress.IndexOf(' '));
            }

            if (CellAddress.IndexOf(':') < 0)
            {
                ret = GetRowColFromAddress(CellAddress, out FromRow, out FromColumn, out fixedFromRow, out fixedFromColumn);
                ToColumn = FromColumn;
                ToRow = FromRow;
                fixedToRow = fixedFromRow;
                fixedToColumn = fixedFromColumn;
            }
            else
            {
                string[] cells = CellAddress.Split(':');
                ret = GetRowColFromAddress(cells[0], out FromRow, out FromColumn, out fixedFromRow, out fixedFromColumn);
                if (ret)
                    ret = GetRowColFromAddress(cells[1], out ToRow, out ToColumn, out fixedToRow, out fixedToColumn);
                else
                {
                    GetRowColFromAddress(cells[1], out ToRow, out ToColumn, out fixedToRow, out fixedToColumn);
                }

                if (FromColumn <= 0)
                    FromColumn = 1;
                if (FromRow <= 0)
                    FromRow = 1;
                if (ToColumn <= 0)
                    ToColumn = ExcelPackage.MaxColumns;
                if (ToRow <= 0)
                    ToRow = ExcelPackage.MaxRows;
            }
            return ret;
        }
        /// <summary>
        /// Get the row/column for n Cell-address
        /// </summary>
        /// <param name="CellAddress">The address</param>
        /// <param name="Row">Returns Tthe row</param>
        /// <param name="Column">Returns the column</param>
        /// <returns>true if valid</returns>
        internal static bool GetRowColFromAddress(string CellAddress, out int Row, out int Column)
        {
            return GetRowCol(CellAddress, out Row, out Column, true);
        }
        internal static bool GetRowColFromAddress(string CellAddress, out int row, out int col, out bool fixedRow, out bool fixedCol)
        {
            return GetRowCol(CellAddress, out row, out col, true, out fixedRow, out fixedCol);
        }

        /// <summary>
        /// Get the row/column for a Cell-address
        /// </summary>
        /// <param name="address">the address</param>
        /// <param name="row">returns the row</param>
        /// <param name="col">returns the column</param>
        /// <param name="throwException">throw exception if invalid, otherwise returns false</param>
        /// <returns></returns>
        internal static bool GetRowCol(string address, out int row, out int col, bool throwException)
        {
            bool fixedRow, fixedCol;
            return GetRowCol(address, out row, out col, throwException, out fixedRow, out fixedCol);
        }
        internal static bool GetRowCol(string address, out int row, out int col, bool throwException, out bool fixedRow, out bool fixedCol)
        {
          bool colPart = true;
          int colStartIx = 0;
          int colLength = 0;
          col = 0;
          row = 0;
          fixedRow = false;
          fixedCol = false;

          if (address.EndsWith("#REF!"))
          {
            row = 0;
            col = 0;
            return true;
          }

          int sheetMarkerIndex = address.IndexOf('!');
          if (sheetMarkerIndex >= 0)
          {
            colStartIx = sheetMarkerIndex + 1;
          }
          address = address.ToUpper(CultureInfo.InvariantCulture);
          for (int i = colStartIx; i < address.Length; i++)
          {
            char c = address[i];
            if (colPart && (c >= 'A' && c <= 'Z') && colLength <= 3)
            {
              col *= 26;
              col += ((int)c) - 64;
              colLength++;
            }
            else if (c >= '0' && c <= '9')
            {
              row *= 10;
              row += ((int)c) - 48;
              colPart = false;
            }
            else if (c == '$')
            {
              if (i == colStartIx)
              {
                colStartIx++;
                fixedCol = true;
              }
              else
              {
                colPart = false;
                fixedRow = true;
              }
            }
            else
            {
              row = 0;
              col = 0;
              if (throwException)
              {
                throw (new Exception(string.Format("Invalid Address format {0}", address)));
              }
              else
              {
                return false;
              }
            }
          }
          return row != 0 || col != 0;
        }

        private static int GetColumn(string sCol)
        {
            int col = 0;
            int len = sCol.Length - 1;
            for (int i = len; i >= 0; i--)
            {
                col += (((int)sCol[i]) - 64) * (int)(Math.Pow(26, len - i));
            }
            return col;
        }
        #region GetAddress
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="Row">The number of the row</param>
        /// <param name="Column">The number of the column in the worksheet</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int Row, int Column)
        {
            return GetAddress(Row, Column,false);
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="Row">The number of the row</param>
        /// <param name="Column">The number of the column in the worksheet</param>
        /// <param name="AbsoluteRow">Absolute row</param>
        /// <param name="AbsoluteCol">Absolute column</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int Row, bool AbsoluteRow, int Column, bool AbsoluteCol)
        {
            return ( AbsoluteCol ? "$" : "") + GetColumnLetter(Column) + ( AbsoluteRow ? "$" : "") + Row.ToString();
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="Row">The number of the row</param>
        /// <param name="Column">The number of the column in the worksheet</param>
        /// <param name="Absolute">Get an absolute address ($A$1)</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int Row, int Column, bool Absolute)
        {
            if (Row == 0 || Column == 0)
            {
                return "#REF!";
            }
            if (Absolute)
            {
                return ("$" + GetColumnLetter(Column) + "$" + Row.ToString());
            }
            else
            {
                return (GetColumnLetter(Column) + Row.ToString());
            }
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="FromRow">From row number</param>
        /// <param name="FromColumn">From column number</param>
        /// <param name="ToRow">To row number</param>
        /// <param name="ToColumn">From column number</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn)
        {
            return GetAddress(FromRow, FromColumn, ToRow, ToColumn, false);
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="FromRow">From row number</param>
        /// <param name="FromColumn">From column number</param>
        /// <param name="ToRow">To row number</param>
        /// <param name="ToColumn">From column number</param>
        /// <param name="Absolute">if true address is absolute (like $A$1)</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn, bool Absolute)
        {
            if (FromRow == ToRow && FromColumn == ToColumn)
            {
                return GetAddress(FromRow, FromColumn, Absolute);
            }
            else
            {
                if (FromRow == 1 && ToRow == ExcelPackage.MaxRows)
                {
                    var absChar = Absolute ? "$" : "";
                    return absChar + GetColumnLetter(FromColumn) + ":" + absChar + GetColumnLetter(ToColumn);
                }
                else if(FromColumn==1 && ToColumn==ExcelPackage.MaxColumns)
                {
                    var absChar = Absolute ? "$" : "";
                    return absChar + FromRow.ToString() + ":" + absChar + ToRow.ToString();
                }
                else
                {
                    return GetAddress(FromRow, FromColumn, Absolute) + ":" + GetAddress(ToRow, ToColumn, Absolute);
                }
            }
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="FromRow">From row number</param>
        /// <param name="FromColumn">From column number</param>
        /// <param name="ToRow">To row number</param>
        /// <param name="ToColumn">From column number</param>
        /// <param name="FixedFromColumn"></param>
        /// <param name="FixedFromRow"></param>
        /// <param name="FixedToColumn"></param>
        /// <param name="FixedToRow"></param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn, bool FixedFromRow, bool FixedFromColumn, bool FixedToRow, bool  FixedToColumn)
        {
            if (FromRow == ToRow && FromColumn == ToColumn)
            {
                return GetAddress(FromRow, FixedFromRow, FromColumn, FixedFromColumn);
            }
            else
            {
                if (FromRow == 1 && ToRow == ExcelPackage.MaxRows)
                {
                    return GetColumnLetter(FromColumn, FixedFromColumn) + ":" + GetColumnLetter(ToColumn, FixedToColumn);
                }
                else if (FromColumn == 1 && ToColumn == ExcelPackage.MaxColumns)
                {                    
                    return (FixedFromRow ? "$":"") + FromRow.ToString() + ":" + (FixedToRow ? "$":"") + ToRow.ToString();
                }
                else
                {
                    return GetAddress(FromRow, FixedFromRow, FromColumn, FixedFromColumn) + ":" + GetAddress(ToRow, FixedToRow, ToColumn, FixedToColumn);
                }
            }
        }
        /// <summary>
        /// Get the full address including the worksheet name
        /// </summary>
        /// <param name="worksheetName">The name of the worksheet</param>
        /// <param name="address">The address</param>
        /// <returns>The full address</returns>
        public static string GetFullAddress(string worksheetName, string address)
        {
            return GetFullAddress(worksheetName, address, true);
        }
        internal static string GetFullAddress(string worksheetName, string address, bool fullRowCol)
        {
               if (address.IndexOf("!") == -1 || address=="#REF!")
               {
                   if (fullRowCol)
                   {
                       string[] cells = address.Split(':');
                       if (cells.Length > 0)
                       {
                           address = string.Format("'{0}'!{1}", worksheetName, cells[0]);
                           if (cells.Length > 1)
                           {
                               address += string.Format(":{0}", cells[1]);
                           }
                       }
                   }
                   else
                   {
                       var a = new ExcelAddressBase(address);
                       if ((a._fromRow == 1 && a._toRow == ExcelPackage.MaxRows) || (a._fromCol == 1 && a._toCol == ExcelPackage.MaxColumns))
                       {
                           address = string.Format("'{0}'!{1}{2}:{3}{4}", worksheetName, ExcelAddress.GetColumnLetter(a._fromCol), a._fromRow, ExcelAddress.GetColumnLetter(a._toCol), a._toRow);
                       }
                       else
                       {
                           address=GetFullAddress(worksheetName, address, true);
                       }
                   }
               }
               return address;
        }
        #endregion
        #region IsValidCellAddress
        public static bool IsValidAddress(string address)
        {
            address = address.ToUpper(CultureInfo.InvariantCulture);
            string r1 = "", c1 = "", r2 = "", c2 = "";
            bool isSecond = false;
            for (int i = 0; i < address.Length; i++)
            {
                if (address[i] >= 'A' && address[i] <= 'Z')
                {
                    if (isSecond == false)
                    {
                        if (r1 != "") return false;
                        c1 += address[i];
                        if (c1.Length > 3) return false;
                    }
                    else
                    {
                        if (r2 != "") return false;
                        c2 += address[i];
                        if (c2.Length > 3) return false;
                    }
                }
                else if (address[i] >= '0' && address[i] <= '9')
                {
                    if (isSecond == false)
                    {
                        r1 += address[i];
                        if (r1.Length > 7) return false;
                    }
                    else
                    {
                        r2 += address[i];
                        if (r2.Length > 7) return false;
                    }
                }
                else if (address[i] == ':')
                {
                    isSecond=true;
                }
                else if (address[i] == '$')
                {
                    if (i == address.Length - 1 || address[i + 1] == ':')
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }

            if (r1!="" && c1!="" && r2 == "" && c2 == "")   //Single Cell
            {
                return (GetColumn(c1)<=ExcelPackage.MaxColumns && int.Parse(r1)<=ExcelPackage.MaxRows);   
            }
            else if (r1 != "" && r2 != "" && c1 != "" && c2 != "") //Range
            {
                var iR2 = int.Parse(r2);
                var iC2 = GetColumn(c2);

                return GetColumn(c1) <= iC2 && int.Parse(r1) <= iR2 &&
                    iC2 <= ExcelPackage.MaxColumns && iR2 <= ExcelPackage.MaxRows;
                                                    
            }
            else if (r1 == "" && r2 == "" && c1 != "" && c2 != "") //Full Column
            {                
                var c2n=GetColumn(c2);
                return (GetColumn(c1) <= c2n && c2n <= ExcelPackage.MaxColumns);
            }
            else if (r1 != "" && r2 != "" && c1 == "" && c2 == "")
            {
                var iR2 = int.Parse(r2);

                return int.Parse(r1) <= iR2 && iR2 <= ExcelPackage.MaxRows;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Checks that a cell address (e.g. A5) is valid.
        /// </summary>
        /// <param name="cellAddress">The alphanumeric cell address</param>
        /// <returns>True if the cell address is valid</returns>
        public static bool IsValidCellAddress(string cellAddress)
        {
            bool result = false;
            try
            {
                int row, col;
                if (GetRowColFromAddress(cellAddress, out row, out col))
                {
                    if (row > 0 && col > 0 && row <= ExcelPackage.MaxRows && col <= ExcelPackage.MaxColumns)
                        result = true;
                    else
                        result = false;
                }
            }
            catch { }
            return result;
        }
        #endregion
        #region UpdateFormulaReferences
        /// <summary>
        /// Updates the Excel formula so that all the cellAddresses are incremented by the row and column increments
        /// if they fall after the afterRow and afterColumn.
        /// Supports inserting rows and columns into existing templates.
        /// </summary>
        /// <param name="Formula">The Excel formula</param>
        /// <param name="rowIncrement">The amount to increment the cell reference by</param>
        /// <param name="colIncrement">The amount to increment the cell reference by</param>
        /// <param name="afterRow">Only change rows after this row</param>
        /// <param name="afterColumn">Only change columns after this column</param>
        /// <returns></returns>
        internal static string UpdateFormulaReferences(string Formula, int rowIncrement, int colIncrement, int afterRow, int afterColumn, bool setFixed=false)
        {
            //return Translate(Formula, AddToRowColumnTranslator, afterRow, afterColumn, rowIncrement, colIncrement);
            var d=new Dictionary<string, object>();
            try
            {
                var sct = new SourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty);
                var tokens = sct.Tokenize(Formula);
                String f = "";
                foreach (var t in tokens)
                {
                    if (t.TokenType == TokenType.ExcelAddress)
                    {
                        var a = new ExcelAddressBase(t.Value);                        
                        if (rowIncrement > 0)
                        {
                            a = a.AddRow(afterRow, rowIncrement, setFixed);
                        }
                        else if (rowIncrement < 0)
                        {
                            a = a.DeleteRow(afterRow, -rowIncrement, setFixed);
                        }
                        if (colIncrement > 0)
                        {
                            a = a.AddColumn(afterColumn, colIncrement, setFixed);
                        }
                        else if (colIncrement < 0)
                        {
                            a = a.DeleteColumn(afterColumn, -colIncrement, setFixed);
                        }
                        if (a == null || !a.IsValidRowCol())
                        {
                            f += "#REF!";
                        }
                        else
                        {
                            f += a.Address;
                        }


                    }
                    else
                    {
                        f += t.Value;
                    }
                }
                return f;
            }
            catch //Invalid formula, skip updateing addresses
            {
                return Formula;
            }
        }

        #endregion
        #endregion
        #endregion
    }
}
