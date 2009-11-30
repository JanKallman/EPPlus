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
using System.Text.RegularExpressions;

namespace OfficeOpenXml
{
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
        private delegate string dlgTransl(string part, int row, int col);
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
            return Translate(value, ToAbs, row, col);
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
            return Translate(value, ToR1C1, row, col);
        }
        /// <summary>
        /// Translates betweein R1C1 or absolut addresses
        /// </summary>
        /// <param name="value">The addresss/function</param>
        /// <param name="addressTranslator">The translating function</param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private static string Translate(string value, dlgTransl addressTranslator, int row, int col)
        {
            if (value == "") return "";
            bool isText = false;
            string ret = "";
            string part = "";
            for (int pos = 0; pos < value.Length; pos++)
            {
                char c = value[pos];
                if (c == '"')
                {
                    if (isText == false && part != "")
                    {
                        ret += addressTranslator(part, row, col);
                        part = "";
                    }
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
                       c == '<' || c == '>' || c == '(' || c == ')') &&
                       (pos == 0 || value[pos - 1] != '[')) //Last part to allow for R1C1 style [-x]
                    {
                        ret += addressTranslator(part, row, col) + c;
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
                addressTranslator(part, row, col);
            }
            return ret;
        }
        /// <summary>
        /// Translate to R1C1
        /// </summary>
        /// <param name="part">the value to be translated</param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private static string ToR1C1(string part, int row, int col)
        {
            int addrRow, addrCol;
            string Ret = "R";
            if (GetRowCol(part, out addrRow, out addrCol, false))
            {
                if (part.StartsWith("$"))
                {
                    Ret += addrRow.ToString();
                }
                else if (addrRow - row != 0)
                {
                    Ret += string.Format("[{0}]", addrRow - row);
                }

                if (part.IndexOf('$', 1) > 0)
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
        /// <returns></returns>
        private static string ToAbs(string part, int row, int col)
        {
            string check = part.ToUpper();

            int rStart = check.IndexOf("R");
            if (rStart != 0) return part;
            if (part.Length == 1) //R
            {
                return GetAddress(row, col);
            }

            int cStart = check.IndexOf("C");
            if (cStart == -1)
            {
                int RNum = GetRC(part.Substring(1, cStart), row);
                if (RNum > int.MinValue)
                {
                    return GetAddress(RNum, col); ;
                }
                else
                {
                    return part;
                }
            }
            else
            {
                int RNum = GetRC(part.Substring(1, cStart - 1), row);
                int CNum = GetRC(part.Substring(cStart + 1, part.Length - cStart - 1), col);
                if (RNum > int.MinValue && CNum > int.MinValue)
                {
                    return GetAddress(RNum, CNum);
                }
                else
                {
                    return part;
                }
            }
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
        /// <returns></returns>
        private static int GetRC(string value, int OffsetValue)
        {
            if (value == "") return OffsetValue;
            int num;
            if (value[0] == '[' && value[value.Length - 1] == ']') //Offset?
            {
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

            if (iColumnNumber < 1)
            {
                throw new Exception("Column number is out of range");
            }

            string sCol = "";
            do
            {
                sCol = ((char)('A' + ((iColumnNumber - 1) % 26))) + sCol;
                iColumnNumber = (iColumnNumber - ((iColumnNumber - 1) % 26)) / 26;
            }
            while (iColumnNumber > 0);
            return sCol;
        }
        #endregion
        /// <summary>
        /// Get the row/columns for an address
        /// </summary>
        /// <param name="cellAddress">The address</param>
        /// <param name="_fromCol">Returns the from column</param>
        /// <param name="_fromRow">Returns the to column</param>
        /// <param name="_toCol">Returns the from row</param>
        /// <param name="_toRow">Returns the to row</param>
        internal static void GetRowColFromAddress(string cellAddress, out int _fromRow, out int _fromCol, out int _toRow, out int _toCol)
        {
            cellAddress = cellAddress.ToUpper();
            string[] cells = cellAddress.Split(':');
            GetRowCol(cells[0], out _fromRow, out _fromCol);
            if (cells.Length > 1)
            {
                GetRowCol(cells[1], out _toRow, out _toCol);
            }
            else
            {
                _toCol = _fromCol;
                _toRow = _fromRow;
            }
        }
        /// <summary>
        /// Get the row/column for a Cell-address
        /// </summary>
        /// <param name="address">the address</param>
        /// <param name="row">returns the row</param>
        /// <param name="col">returns the column</param>
        /// <returns>true if valid</returns>
        internal static bool GetRowCol(string address, out int row, out int col)
        {
            return GetRowCol(address, out row, out col, true);
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
            bool colPart = true;
            string sRow = "", sCol = "";
            col = 0;
            for (int i = 0; i < address.Length; i++)
            {
                if ((address[i] >= 'A' && address[i] <= 'Z') && colPart && sCol.Length <= 3)
                {
                    sCol += address[i];
                }
                else if (address[i] >= '0' && address[i] <= '9')
                {
                    sRow += address[i];
                    colPart = false;
                }
                else if (address[i] != '$') // $ is ignored here
                {
                    if (throwException)
                    {
                        throw (new Exception(string.Format("Invalid Address format {0}", address)));
                    }
                    else
                    {
                        row = 0;
                        col = 0;
                        return false;
                    }
                }
            }
            // Get the column number
            if (sCol != "")
            {
                int len = sCol.Length - 1;
                for (int i = len; i >= 0; i--)
                {
                    col += (((int)sCol[i]) - 64) * (int)(Math.Pow(26, len - i));
                }
            }
            // Get the row number
            if (sRow == "")
            {
                if (throwException)
                {
                    throw (new Exception(string.Format("Invalid Address format {0}", address)));
                }
                else
                {
                    row = 0;
                    return false;
                }
            }
            else
            {
                int.TryParse(sRow, out row);
            }
            return true;
        }
        #region GetAddress
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="iRow">The number of the row</param>
        /// <param name="iColumn">The number of the column in the worksheet</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int Row, int Column)
        {
            return (GetColumnLetter(Column) + Row.ToString());
        }
        /// <summary>
        /// Returns the AlphaNumeric representation that Excel expects for a Cell Address
        /// </summary>
        /// <param name="iRow">The number of the row</param>
        /// <param name="iColumn">The number of the column in the worksheet</param>
        /// <returns>The cell address in the format A1</returns>
        public static string GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn)
        {
            return GetColumnLetter(FromColumn) + FromRow.ToString() + ":" + GetColumnLetter(ToColumn) + ToRow.ToString();
        }
        #endregion

        #region IsValidCellAddress
        /// <summary>
        /// Checks that a cell address (e.g. A5) is valid.
        /// </summary>
        /// <param name="cellAddress">The alphanumeric cell address</param>
        /// <returns>True if the cell address is valid</returns>
        public static bool IsValidCellAddress(string cellAddress)
        {
            int row, col;
            GetRowCol(cellAddress, out row, out col);

            if (GetAddress(row, col) == cellAddress)
                return (true);
            else
                return (false);
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
        public static string UpdateFormulaReferences(string Formula, int rowIncrement, int colIncrement, int afterRow, int afterColumn)
        {
            string newFormula = "";

            Regex getAlphaNumeric = new Regex(@"[^a-zA-Z0-9]", RegexOptions.IgnoreCase);
            Regex getSigns = new Regex(@"[a-zA-Z0-9]", RegexOptions.IgnoreCase);

            string alphaNumeric = getAlphaNumeric.Replace(Formula, " ").Replace("  ", " ");
            string signs = getSigns.Replace(Formula, " ");
            char[] chrSigns = signs.ToCharArray();
            int count = 0;
            int length = 0;
            foreach (string cellAddress in alphaNumeric.Split(' '))
            {
                count++;
                length += cellAddress.Length;

                // if the cellAddress contains an alpha part followed by a number part, then it is a valid cellAddress

                int row, col;
                GetRowCol(cellAddress, out row, out col);
                string newCellAddress = "";
                if (GetAddress(row, col) == cellAddress)   // this checks if the cellAddress is valid
                {
                    // we have a valid cell address so change its value (if necessary)
                    if (row >= afterRow)
                        row += rowIncrement;
                    if (col >= afterColumn)
                        col += colIncrement;
                    newCellAddress = GetAddress(row, col);
                }
                if (newCellAddress == "")
                {
                    newFormula += cellAddress;
                }
                else
                {
                    newFormula += newCellAddress;
                }

                for (int i = length; i < signs.Length; i++)
                {
                    if (chrSigns[i] == ' ')
                        break;
                    if (chrSigns[i] != ' ')
                    {
                        length++;
                        newFormula += chrSigns[i].ToString();
                    }
                }
            }
            return (newFormula);
        }

        #endregion
        #endregion
        #endregion
    }
}
