/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
 * ******************************************************************************
 * Jan Källman		    Initial Release		        2010-01-28
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 * Eyal Seagull		    Conditional Formatting      2012-04-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Data;
using System.Threading;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.Style;
using System.Xml;
using System.Drawing;
using System.Globalization;
using System.Collections;
using OfficeOpenXml.Table;
using System.Text.RegularExpressions;
using System.IO;
using System.Linq;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using System.Reflection;
using OfficeOpenXml.Style.XmlAccess;
using System.Security;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using w = System.Windows;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Compatibility;

namespace OfficeOpenXml
{
    /// <summary>
    /// A range of cells 
    /// </summary>
    public class ExcelRangeBase : ExcelAddress, IExcelCell, IDisposable, IEnumerable<ExcelRangeBase>, IEnumerator<ExcelRangeBase>
    {
        /// <summary>
        /// Reference to the worksheet
        /// </summary>
        protected ExcelWorksheet _worksheet;
        internal ExcelWorkbook _workbook = null;
        private delegate void _changeProp(ExcelRangeBase range, _setValue method, object value);
        private delegate void _setValue(ExcelRangeBase range, object value, int row, int col);
        private _changeProp _changePropMethod;
        private int _styleID;
        private class CopiedCell
        {
            internal int Row { get; set; }
            internal int Column { get; set; }
            internal object Value { get; set; }
            internal string Type { get; set; }
            internal object Formula { get; set; }
            internal int? StyleID { get; set; }
            internal Uri HyperLink { get; set; }
            internal ExcelComment Comment { get; set; }
            internal Byte Flag { get; set; }
        }
        #region Constructors
        internal ExcelRangeBase(ExcelWorksheet xlWorksheet)
        {
            _worksheet = xlWorksheet;
            _ws = _worksheet.Name;
            _workbook = _worksheet.Workbook;
            SetDelegate();
        }
        /// <summary>
        /// On change address handler
        /// </summary>
        protected internal override void ChangeAddress()
        {
            if (Table != null)
            {
                SetRCFromTable(_workbook._package, null);
            }
            SetDelegate();
        }
        internal ExcelRangeBase(ExcelWorksheet xlWorksheet, string address) :
            base(xlWorksheet == null ? "" : xlWorksheet.Name, address)
        {
            _worksheet = xlWorksheet;
            _workbook = _worksheet.Workbook;
            base.SetRCFromTable(_worksheet._package, null);
            if (string.IsNullOrEmpty(_ws)) _ws = _worksheet == null ? "" : _worksheet.Name;
            SetDelegate();
        }
        internal ExcelRangeBase(ExcelWorkbook wb, ExcelWorksheet xlWorksheet, string address, bool isName) :
            base(xlWorksheet == null ? "" : xlWorksheet.Name, address, isName)
        {
            SetRCFromTable(wb._package, null);
            _worksheet = xlWorksheet;
            _workbook = wb;
            if (string.IsNullOrEmpty(_ws)) _ws = (xlWorksheet == null ? null : xlWorksheet.Name);
            SetDelegate();
        }
        #endregion
        #region Set Value Delegates        
        private static _changeProp _setUnknownProp = SetUnknown;
        private static _changeProp _setSingleProp = SetSingle;
        private static _changeProp _setRangeProp = SetRange;
        private static _changeProp _setMultiProp = SetMultiRange;
        private void SetDelegate()
        {
            if (_fromRow == -1)
            {
                _changePropMethod = SetUnknown;
            }
            //Single cell
            else if (_fromRow == _toRow && _fromCol == _toCol && Addresses == null)
            {
                _changePropMethod = SetSingle;
            }
            //Range (ex A1:A2)
            else if (Addresses == null)
            {
                _changePropMethod = SetRange;
            }
            //Multi Range (ex A1:A2,C1:C2)
            else
            {
                _changePropMethod = SetMultiRange;
            }
        }
        /// <summary>
        /// We dont know the address yet. Set the delegate first time a property is set.
        /// </summary>
        /// <param name="range"></param>
        /// <param name="valueMethod"></param>
        /// <param name="value"></param>
        private static void SetUnknown(ExcelRangeBase range, _setValue valueMethod, object value)
        {
            //Address is not set use, selected range
            if (range._fromRow == -1)
            {
                range.SetToSelectedRange();
            }
            range.SetDelegate();
            range._changePropMethod(range, valueMethod, value);
        }
        /// <summary>
        /// Set a single cell
        /// </summary>
        /// <param name="range"></param>
        /// <param name="valueMethod"></param>
        /// <param name="value"></param>
        private static void SetSingle(ExcelRangeBase range, _setValue valueMethod, object value)
        {
            valueMethod(range, value, range._fromRow, range._fromCol);
        }
        /// <summary>
        /// Set a range
        /// </summary>
        /// <param name="range"></param>
        /// <param name="valueMethod"></param>
        /// <param name="value"></param>
        private static void SetRange(ExcelRangeBase range, _setValue valueMethod, object value)
        {
            range.SetValueAddress(range, valueMethod, value);
        }
        /// <summary>
        /// Set a multirange (A1:A2,C1:C2)
        /// </summary>
        /// <param name="range"></param>
        /// <param name="valueMethod"></param>
        /// <param name="value"></param>
        private static void SetMultiRange(ExcelRangeBase range, _setValue valueMethod, object value)
        {
            range.SetValueAddress(range, valueMethod, value);
            foreach (var address in range.Addresses)
            {
                range.SetValueAddress(address, valueMethod, value);
            }
        }
        /// <summary>
        /// Set the property for an address
        /// </summary>
        /// <param name="address"></param>
        /// <param name="valueMethod"></param>
        /// <param name="value"></param>
        private void SetValueAddress(ExcelAddress address, _setValue valueMethod, object value)
        {
            IsRangeValid("");
            if (_fromRow == 1 && _fromCol == 1 && _toRow == ExcelPackage.MaxRows && _toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
            {
                throw (new ArgumentException("Can't reference all cells. Please use the indexer to set the range"));
            }
            else
            {
                if (value is object[,] && valueMethod == Set_Value)
                {
                    // only simple set value is supported for bulk copy
                    _worksheet.SetRangeValueInner(address.Start.Row, address.Start.Column, address.End.Row, address.End.Column, (object[,])value);
                }
                else
                {
                    for (int col = address.Start.Column; col <= address.End.Column; col++)
                    {
                        for (int row = address.Start.Row; row <= address.End.Row; row++)
                        {
                            valueMethod(this, value, row, col);
                        }
                    }
                }
            }
        }
        #endregion
        #region Set property methods
        private static _setValue _setStyleIdDelegate = Set_StyleID;
        private static _setValue _setValueDelegate = Set_Value;
        private static _setValue _setHyperLinkDelegate = Set_HyperLink;
        private static _setValue _setIsRichTextDelegate = Set_IsRichText;
        private static _setValue _setExistsCommentDelegate = Exists_Comment;
        private static _setValue _setCommentDelegate = Set_Comment;

        private static void Set_StyleID(ExcelRangeBase range, object value, int row, int col)
        {
            range._worksheet.SetStyleInner(row, col, (int)value);
        }
        private static void Set_StyleName(ExcelRangeBase range, object value, int row, int col)
        {
            range._worksheet.SetStyleInner(row, col, range._styleID);
        }
        private static void Set_Value(ExcelRangeBase range, object value, int row, int col)
        {
            var sfi = range._worksheet._formulas.GetValue(row, col);
            if (sfi is int)
            {
                range.SplitFormulas(range._worksheet.Cells[row, col]);
            }
            if (sfi != null) range._worksheet._formulas.SetValue(row, col, string.Empty);
            range._worksheet.SetValueInner(row, col, value);
        }
        private static void Set_Formula(ExcelRangeBase range, object value, int row, int col)
        {
            var f = range._worksheet._formulas.GetValue(row, col);
            if (f is int && (int)f >= 0) range.SplitFormulas(range._worksheet.Cells[row, col]);

            string formula = (value == null ? string.Empty : value.ToString());
            if (formula == string.Empty)
            {
                range._worksheet._formulas.SetValue(row, col, string.Empty);
            }
            else
            {
                if (formula[0] == '=') value = formula.Substring(1, formula.Length - 1); // remove any starting equalsign.
                range._worksheet._formulas.SetValue(row, col, formula);
                range._worksheet.SetValueInner(row, col, null);
            }
        }
        /// <summary>
        /// Handles shared formulas
        /// </summary>
        /// <param name="range">The range</param>
        /// <param name="value">The  formula</param>
        /// <param name="address">The address of the formula</param>
        /// <param name="IsArray">If the forumla is an array formula.</param>
        private static void Set_SharedFormula(ExcelRangeBase range, string value, ExcelAddress address, bool IsArray)
        {
            if (range._fromRow == 1 && range._fromCol == 1 && range._toRow == ExcelPackage.MaxRows && range._toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
            {
                throw (new InvalidOperationException("Can't set a formula for the entire worksheet"));
            }
            else if (address.Start.Row == address.End.Row && address.Start.Column == address.End.Column && !IsArray)             //is it really a shared formula? Arrayformulas can be one cell only
            {
                //Nope, single cell. Set the formula
                Set_Formula(range, value, address.Start.Row, address.Start.Column);
                return;
            }
            range.CheckAndSplitSharedFormula(address);
            ExcelWorksheet.Formulas f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
            f.Formula = value;
            f.Index = range._worksheet.GetMaxShareFunctionIndex(IsArray);
            f.Address = address.FirstAddress;
            f.StartCol = address.Start.Column;
            f.StartRow = address.Start.Row;
            f.IsArray = IsArray;

            range._worksheet._sharedFormulas.Add(f.Index, f);

            for (int col = address.Start.Column; col <= address.End.Column; col++)
            {
                for (int row = address.Start.Row; row <= address.End.Row; row++)
                {
                    range._worksheet._formulas.SetValue(row, col, f.Index);
                    range._worksheet._flags.SetFlagValue(row, col, true, CellFlags.ArrayFormula);
                    range._worksheet.SetValueInner(row, col, null);
                }
            }
        }
        private static void Set_HyperLink(ExcelRangeBase range, object value, int row, int col)
        {
            if (value is Uri)
            {
                range._worksheet._hyperLinks.SetValue(row, col, (Uri)value);

                if (value is ExcelHyperLink)
                {
                    range._worksheet.SetValueInner(row, col, ((ExcelHyperLink)value).Display);
                }
                else
                {
                    var v = range._worksheet.GetValueInner(row, col);
                    if (v == null || v.ToString() == "")
                    {
                        range._worksheet.SetValueInner(row, col, ((Uri)value).OriginalString);
                    }
                }
            }
            else
            {
                range._worksheet._hyperLinks.SetValue(row, col, (Uri)null);
                range._worksheet.SetValueInner(row, col, (Uri)null);
            }
        }
        private static void Set_IsRichText(ExcelRangeBase range, object value, int row, int col)
        {
            range._worksheet._flags.SetFlagValue(row, col, (bool)value, CellFlags.RichText);
        }
        private static void Exists_Comment(ExcelRangeBase range, object value, int row, int col)
        {
            if (range._worksheet._commentsStore.Exists(row, col))
            {
                throw (new InvalidOperationException(string.Format("Cell {0} already contain a comment.", new ExcelCellAddress(row, col).Address)));
            }

        }
        private static void Set_Comment(ExcelRangeBase range, object value, int row, int col)
        {
            string[] v = (string[])value;
            range._worksheet.Comments.Add(new ExcelRangeBase(range._worksheet, GetAddress(range._fromRow, range._fromCol)), v[0], v[1]);
        }
        #endregion
        private void SetToSelectedRange()
        {
            if (_worksheet.View.SelectedRange == "")
            {
                Address = "A1";
            }
            else
            {
                Address = _worksheet.View.SelectedRange;
            }
        }
        private void IsRangeValid(string type)
        {
            if (_fromRow <= 0)
            {
                if (_address == "")
                {
                    SetToSelectedRange();
                }
                else
                {
                    if (type == "")
                    {
                        throw (new InvalidOperationException(string.Format("Range is not valid for this operation: {0}", _address)));
                    }
                    else
                    {
                        throw (new InvalidOperationException(string.Format("Range is not valid for {0} : {1}", type, _address)));
                    }
                }
            }
        }
        internal void UpdateAddress(string address)
        {
            throw new NotImplementedException();
        }

        #region Public Properties
        /// <summary>
        /// The styleobject for the range.
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                IsRangeValid("styling");
                int s = 0;
                if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref s)) //Cell exists
                {
                    if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref s)) //No, check Row style
                    {
                        var c = Worksheet.GetColumn(_fromCol);
                        if (c == null)
                        {
                            s = 0;
                        }
                        else
                        {
                            s = c.StyleID;
                        }
                    }
                }
                return _worksheet.Workbook.Styles.GetStyleObject(s, _worksheet.PositionID, Address);
            }
        }
        /// <summary>
        /// The named style
        /// </summary>
        public string StyleName
        {
            get
            {
                IsRangeValid("styling");
                int xfId;
                if (_fromRow == 1 && _toRow == ExcelPackage.MaxRows)
                {
                    xfId = GetColumnStyle(_fromCol);
                }
                else if (_fromCol == 1 && _toCol == ExcelPackage.MaxColumns)
                {
                    xfId = 0;
                    if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref xfId))
                    {
                        xfId = GetColumnStyle(_fromCol);
                    }
                }
                else
                {
                    xfId = 0;
                    if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref xfId))
                    {
                        if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref xfId))
                        {
                            xfId = GetColumnStyle(_fromCol);
                        }
                    }
                }
                int nsID;
                if (xfId <= 0)
                {
                    nsID = Style.Styles.CellXfs[0].XfId;
                }
                else
                {
                    nsID = Style.Styles.CellXfs[xfId].XfId;
                }
                foreach (var ns in Style.Styles.NamedStyles)
                {
                    if (ns.StyleXfId == nsID)
                    {
                        return ns.Name;
                    }
                }

                return "";
            }
            set
            {
                _styleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
                int col = _fromCol;
                if (_fromRow == 1 && _toRow == ExcelPackage.MaxRows)    //Full column
                {
                    ExcelColumn column;
                    var c = _worksheet.GetValue(0, _fromCol);
                    if (c == null)
                    {
                        column = _worksheet.Column(_fromCol);
                    }
                    else
                    {
                        column = (ExcelColumn)c;
                    }

                    column.StyleName = value;
                    column.StyleID = _styleID;

                    var cols = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 0, _fromCol + 1, 0, _toCol);
                    if (cols.Next())
                    {
                        col = _fromCol;
                        while (column.ColumnMin <= _toCol)
                        {
                            if (column.ColumnMax > _toCol)
                            {
                                var newCol = _worksheet.CopyColumn(column, _toCol + 1, column.ColumnMax);
                                column.ColumnMax = _toCol;
                            }

                            column._styleName = value;
                            column.StyleID = _styleID;

                            if (cols.Value._value == null)
                            {
                                break;
                            }
                            else
                            {
                                var nextCol = (ExcelColumn)cols.Value._value;
                                if (column.ColumnMax < nextCol.ColumnMax - 1)
                                {
                                    column.ColumnMax = nextCol.ColumnMax - 1;
                                }
                                column = nextCol;
                                cols.Next();
                            }
                        }
                    }
                    if (column.ColumnMax < _toCol)
                    {
                        column.ColumnMax = _toCol;
                    }

                    if (_fromCol == 1 && _toCol == ExcelPackage.MaxColumns) //FullRow
                    {
                        var rows = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 1, 0, ExcelPackage.MaxRows, 0);
                        rows.Next();
                        while (rows.Value._value != null)
                        {
                            _worksheet.SetStyleInner(rows.Row, 0, _styleID);
                            if (!rows.Next())
                            {
                                break;
                            }
                        }
                    }
                }
                else if (_fromCol == 1 && _toCol == ExcelPackage.MaxColumns) //FullRow
                {
                    for (int r = _fromRow; r <= _toRow; r++)
                    {
                        _worksheet.Row(r)._styleName = value;
                        _worksheet.Row(r).StyleID = _styleID;
                    }
                }

                if (!((_fromRow == 1 && _toRow == ExcelPackage.MaxRows) || (_fromCol == 1 && _toCol == ExcelPackage.MaxColumns))) //Cell specific
                {
                    for (int c = _fromCol; c <= _toCol; c++)
                    {
                        for (int r = _fromRow; r <= _toRow; r++)
                        {
                            _worksheet.SetStyleInner(r, c, _styleID);
                        }
                    }
                }
                else //Only set name on created cells. (uncreated cells is set on full row or full column).
                {
                    var cells = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
                    while (cells.Next())
                    {
                        _worksheet.SetStyleInner(cells.Row, cells.Column, _styleID);
                    }
                }
            }
        }

        private int GetColumnStyle(int col)
        {
            object c = null;
            if (_worksheet.ExistsValueInner(0, col, ref c))
            {
                return (c as ExcelColumn).StyleID;
            }
            else
            {
                int row = 0;
                if (_worksheet._values.PrevCell(ref row, ref col))
                {
                    var column = _worksheet.GetValueInner(row, col) as ExcelColumn;
                    if (column.ColumnMax >= col)
                    {
                        return _worksheet.GetStyleInner(row, col);
                    }
                }
            }
            return 0;
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
                int s = 0;
                if (!_worksheet.ExistsStyleInner(_fromRow, _fromCol, ref s))
                {
                    if (!_worksheet.ExistsStyleInner(_fromRow, 0, ref s))
                    {
                        s = _worksheet.GetStyleInner(0, _fromCol);
                    }
                }
                return s;
            }
            set
            {
                _changePropMethod(this, _setStyleIdDelegate, value);
            }
        }
        /// <summary>
        /// Set the range to a specific value
        /// </summary>
        public object Value
        {
            get
            {
                if (IsName)
                {
                    if (_worksheet == null)
                    {
                        return _workbook._names[_address].NameValue;
                    }
                    else
                    {
                        return _worksheet.Names[_address].NameValue;
                    }
                }
                else
                {
                    if (_fromRow == _toRow && _fromCol == _toCol)
                    {
                        return _worksheet.GetValue(_fromRow, _fromCol);
                    }
                    else
                    {
                        return GetValueArray();
                    }
                }
            }
            set
            {
                if (IsName)
                {
                    if (_worksheet == null)
                    {
                        _workbook._names[_address].NameValue = value;
                    }
                    else
                    {
                        _worksheet.Names[_address].NameValue = value;
                    }
                }
                else
                {
                    _changePropMethod(this, _setValueDelegate, value);
                }
            }
        }

        private bool IsInfinityValue(object value)
        {
            double? valueAsDouble = value as double?;

            if (valueAsDouble.HasValue &&
                (double.IsNegativeInfinity(valueAsDouble.Value) || double.IsPositiveInfinity(valueAsDouble.Value)))
            {
                return true;
            }

            return false;
        }

        private object GetValueArray()
        {
            ExcelAddressBase addr;
            if (_fromRow == 1 && _fromCol == 1 && _toRow == ExcelPackage.MaxRows && _toCol == ExcelPackage.MaxColumns)
            {
                addr = _worksheet.Dimension;
                if (addr == null) return null;
            }
            else
            {
                addr = this;
            }
            object[,] v = new object[addr._toRow - addr._fromRow + 1, addr._toCol - addr._fromCol + 1];

            for (int col = addr._fromCol; col <= addr._toCol; col++)
            {
                for (int row = addr._fromRow; row <= addr._toRow; row++)
                {
                    object o = null;
                    if (_worksheet.ExistsValueInner(row, col, ref o))
                    {
                        if (_worksheet._flags.GetFlagValue(row, col, CellFlags.RichText))
                        {
                            v[row - addr._fromRow, col - addr._fromCol] = GetRichText(row, col).Text;
                        }
                        else
                        {
                            v[row - addr._fromRow, col - addr._fromCol] = o;
                        }
                    }
                }
            }
            return v;
        }
        private ExcelAddressBase GetAddressDim(ExcelRangeBase addr)
        {
            int fromRow, fromCol, toRow, toCol;
            var d = _worksheet.Dimension;
            fromRow = addr._fromRow < d._fromRow ? d._fromRow : addr._fromRow;
            fromCol = addr._fromCol < d._fromCol ? d._fromCol : addr._fromCol;

            toRow = addr._toRow > d._toRow ? d._toRow : addr._toRow;
            toCol = addr._toCol > d._toCol ? d._toCol : addr._toCol;

            if (addr._fromRow == fromRow && addr._fromCol == fromCol && addr._toRow == toRow && addr._toCol == _toCol)
            {
                return addr;
            }
            else
            {
                if (_fromRow > _toRow || _fromCol > _toCol)
                {
                    return null;
                }
                else
                {
                    return new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
                }
            }
        }

        private object GetSingleValue()
        {
            if (IsRichText)
            {
                return RichText.Text;
            }
            else
            {
                return _worksheet.GetValueInner(_fromRow, _fromCol);
            }
        }
        /// <summary>
        /// Returns the formatted value.
        /// </summary>
        public string Text
        {
            get
            {
                return GetFormattedText(false);
            }
        }
        /// <summary>
        /// Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property.
        /// Note: Cells containing formulas must be calculated before autofit is called.
        /// Wrapped and merged cells are also ignored.
        /// </summary>
        public void AutoFitColumns()
        {
            AutoFitColumns(_worksheet.DefaultColWidth);
        }

        /// <summary>
        /// Set the column width from the content of the range.
        /// Note: Cells containing formulas are ignored if no calculation is made.
        ///       Wrapped and merged cells are also ignored.
        /// </summary>
        /// <remarks>This method will not work if you run in an environment that does not support GDI</remarks>
        /// <param name="MinimumWidth">Minimum column width</param>
        public void AutoFitColumns(double MinimumWidth)
        {
            AutoFitColumns(MinimumWidth, double.MaxValue);
        }

        /// <summary>
        /// Set the column width from the content of the range.
        /// Note: Cells containing formulas are ignored if no calculation is made.
        ///       Wrapped and merged cells are also ignored.
        ///      Hidden columns are left hidden.
        /// </summary>
        /// <param name="MinimumWidth">Minimum column width</param>
        /// <param name="MaximumWidth">Maximum column width</param>
        public void AutoFitColumns(double MinimumWidth, double MaximumWidth)
        {
            if (_worksheet.Dimension == null)
            {
                return;
            }
            if (_fromCol < 1 || _fromRow < 1)
            {
                SetToSelectedRange();
            }
            var fontCache = new Dictionary<int, Font>();

            bool doAdjust = _worksheet._package.DoAdjustDrawings;
            _worksheet._package.DoAdjustDrawings = false;
            var drawWidths = _worksheet.Drawings.GetDrawingWidths();

            var fromCol = _fromCol > _worksheet.Dimension._fromCol ? _fromCol : _worksheet.Dimension._fromCol;
            var toCol = _toCol < _worksheet.Dimension._toCol ? _toCol : _worksheet.Dimension._toCol;

            if (fromCol > toCol) return; //Issue 15383

            if (Addresses == null)
            {
                SetMinWidth(MinimumWidth, fromCol, toCol);
            }
            else
            {
                foreach (var addr in Addresses)
                {
                    fromCol = addr._fromCol > _worksheet.Dimension._fromCol ? addr._fromCol : _worksheet.Dimension._fromCol;
                    toCol = addr._toCol < _worksheet.Dimension._toCol ? addr._toCol : _worksheet.Dimension._toCol;
                    SetMinWidth(MinimumWidth, fromCol, toCol);
                }
            }

            //Get any autofilter to widen these columns
            var afAddr = new List<ExcelAddressBase>();
            if (_worksheet.AutoFilterAddress != null)
            {
                afAddr.Add(new ExcelAddressBase(_worksheet.AutoFilterAddress._fromRow,
                                                    _worksheet.AutoFilterAddress._fromCol,
                                                    _worksheet.AutoFilterAddress._fromRow,
                                                    _worksheet.AutoFilterAddress._toCol));
                afAddr[afAddr.Count - 1]._ws = WorkSheet;
            }
            foreach (var tbl in _worksheet.Tables)
            {
                if (tbl.AutoFilterAddress != null)
                {
                    afAddr.Add(new ExcelAddressBase(tbl.AutoFilterAddress._fromRow,
                                                                            tbl.AutoFilterAddress._fromCol,
                                                                            tbl.AutoFilterAddress._fromRow,
                                                                            tbl.AutoFilterAddress._toCol));
                    afAddr[afAddr.Count - 1]._ws = WorkSheet;
                }
            }

            var styles = _worksheet.Workbook.Styles;
            var nf = styles.Fonts[styles.CellXfs[0].FontId];
            var fs = FontStyle.Regular;
            if (nf.Bold) fs |= FontStyle.Bold;
            if (nf.UnderLine) fs |= FontStyle.Underline;
            if (nf.Italic) fs |= FontStyle.Italic;
            if (nf.Strike) fs |= FontStyle.Strikeout;
            var nfont = new Font(nf.Name, nf.Size, fs);

            var normalSize = Convert.ToSingle(ExcelWorkbook.GetWidthPixels(nf.Name, nf.Size));

            Bitmap b;
            Graphics g = null;
            try
            {
                //Check for missing GDI+, then use WPF istead.
                b = new Bitmap(1, 1);
                g = Graphics.FromImage(b);
                g.PageUnit = GraphicsUnit.Pixel;
            }
            catch
            {
                return;
            }

            foreach (var cell in this)
            {
                if (_worksheet.Column(cell.Start.Column).Hidden)    //Issue 15338
                    continue;

                if (cell.Merge == true || cell.Style.WrapText) continue;
                var fntID = styles.CellXfs[cell.StyleID].FontId;
                Font f;
                if (fontCache.ContainsKey(fntID))
                {
                    f = fontCache[fntID];
                }
                else
                {
                    var fnt = styles.Fonts[fntID];
                    fs = FontStyle.Regular;
                    if (fnt.Bold) fs |= FontStyle.Bold;
                    if (fnt.UnderLine) fs |= FontStyle.Underline;
                    if (fnt.Italic) fs |= FontStyle.Italic;
                    if (fnt.Strike) fs |= FontStyle.Strikeout;
                    f = new Font(fnt.Name, fnt.Size, fs);

                    fontCache.Add(fntID, f);
                }
                var ind = styles.CellXfs[cell.StyleID].Indent;
                var textForWidth = cell.TextForWidth;
                var t = textForWidth + (ind > 0 && !string.IsNullOrEmpty(textForWidth) ? new string('_', ind) : "");
                if (t.Length > 32000) t = t.Substring(0, 32000); //Issue
                var size = g.MeasureString(t, f, 10000, StringFormat.GenericDefault);

                double width;
                double r = styles.CellXfs[cell.StyleID].TextRotation;
                if (r <= 0)
                {
                    width = (size.Width + 5) / normalSize;
                }
                else
                {
                    r = (r <= 90 ? r : r - 90);
                    width = (((size.Width - size.Height) * Math.Abs(System.Math.Cos(System.Math.PI * r / 180.0)) + size.Height) + 5) / normalSize;
                }

                foreach (var a in afAddr)
                {
                    if (a.Collide(cell) != eAddressCollition.No)
                    {
                        width += 2.25;
                        break;
                    }
                }

                if (width > _worksheet.Column(cell._fromCol).Width)
                {
                    _worksheet.Column(cell._fromCol).Width = width > MaximumWidth ? MaximumWidth : width;
                }
            }
            _worksheet.Drawings.AdjustWidth(drawWidths);
            _worksheet._package.DoAdjustDrawings = doAdjust;
        }

        private void SetMinWidth(double minimumWidth, int fromCol, int toCol)
        {
            var iterator = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, 0, fromCol, 0, toCol);
            var prevCol = fromCol;
            foreach (ExcelCoreValue val in iterator)
            {
                var col = (ExcelColumn)val._value;
                col.Width = minimumWidth;
                if (_worksheet.DefaultColWidth > minimumWidth && col.ColumnMin > prevCol)
                {
                    var newCol = _worksheet.Column(prevCol);
                    newCol.ColumnMax = col.ColumnMin - 1;
                    newCol.Width = minimumWidth;
                }
                prevCol = col.ColumnMax + 1;
            }
            if (_worksheet.DefaultColWidth > minimumWidth && prevCol < toCol)
            {
                var newCol = _worksheet.Column(prevCol);
                newCol.ColumnMax = toCol;
                newCol.Width = minimumWidth;
            }
        }

        internal string TextForWidth
        {
            get
            {
                return GetFormattedText(true);
            }
        }
        private string GetFormattedText(bool forWidthCalc)
        {
            object v = Value;
            if (v == null) return "";
            var styles = Worksheet.Workbook.Styles;
            var nfID = styles.CellXfs[StyleID].NumberFormatId;
            ExcelNumberFormatXml.ExcelFormatTranslator nf = null;
            for (int i = 0; i < styles.NumberFormats.Count; i++)
            {
                if (nfID == styles.NumberFormats[i].NumFmtId)
                {
                    nf = styles.NumberFormats[i].FormatTranslator;
                    break;
                }
            }
            if (nf == null)
            {
                nf = styles.NumberFormats[0].FormatTranslator;  //nf should never be null. If so set to General, Issue 173
            }

            string format, textFormat;
            if (forWidthCalc)
            {
                format = nf.NetFormatForWidth;
                textFormat = nf.NetTextFormatForWidth;
            }
            else
            {
                format = nf.NetFormat;
                textFormat = nf.NetTextFormat;
            }

            return FormatValue(v, nf, format, textFormat);
        }

        internal static string FormatValue(object v, ExcelNumberFormatXml.ExcelFormatTranslator nf, string format, string textFormat)
        {
            if (v is decimal || TypeCompat.IsPrimitive(v))
            {
                double d;
                try
                {
                    d = Convert.ToDouble(v);
                }
                catch
                {
                    return "";
                }

                if (nf.DataType == ExcelNumberFormatXml.eFormatType.Number)
                {
                    if (string.IsNullOrEmpty(nf.FractionFormat))
                    {
                        return d.ToString(format, nf.Culture);
                    }
                    else
                    {
                        return nf.FormatFraction(d);
                    }
                }
                else if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
                {
                    var date = DateTime.FromOADate(d);
                    return GetDateText(date, format, nf);
                }
            }
            else if (v is DateTime)
            {
                if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
                {
                    return GetDateText((DateTime)v, format, nf);
                }
                else
                {
                    double d = ((DateTime)v).ToOADate();
                    if (string.IsNullOrEmpty(nf.FractionFormat))
                    {
                        return d.ToString(format, nf.Culture);
                    }
                    else
                    {
                        return nf.FormatFraction(d);
                    }
                }
            }
            else if (v is TimeSpan)
            {
                if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
                {
                    return GetDateText(new DateTime(((TimeSpan)v).Ticks), format, nf);
                }
                else
                {
                    double d = new DateTime(0).Add((TimeSpan)v).ToOADate();
                    if (string.IsNullOrEmpty(nf.FractionFormat))
                    {
                        return d.ToString(format, nf.Culture);
                    }
                    else
                    {
                        return nf.FormatFraction(d);
                    }
                }
            }
            else
            {
                if (textFormat == "")
                {
                    return v.ToString();
                }
                else
                {
                    return string.Format(textFormat, v);
                }
            }
            return v.ToString();
        }

        private static string GetDateText(DateTime d, string format, ExcelNumberFormatXml.ExcelFormatTranslator nf)
        {
            if(nf.SpecialDateFormat==ExcelNumberFormatXml.ExcelFormatTranslator.eSystemDateFormat.SystemLongDate)
            {
                return d.ToLongDateString();
            }
            else if(nf.SpecialDateFormat == ExcelNumberFormatXml.ExcelFormatTranslator.eSystemDateFormat.SystemLongTime)
            {
                return d.ToLongTimeString();
            }
            else if (nf.SpecialDateFormat == ExcelNumberFormatXml.ExcelFormatTranslator.eSystemDateFormat.SystemShortDate)
            {
                return d.ToShortDateString();
            }
            if (format == "d" || format == "D")
            {
                return d.Day.ToString();
            }
            else if (format == "M")
            {
                return d.Month.ToString();
            }
            else if (format == "m")
            {
                return d.Minute.ToString();
            }
            else if (format.ToLower() == "y" || format.ToLower() == "yy")
            {
                return d.ToString("yy", nf.Culture);
            }
            else if (format.ToLower() == "yyy" || format.ToLower() == "yyyy")
            {
                return d.ToString("yyy", nf.Culture);
            }
            else
            {
                return d.ToString(format, nf.Culture);
            }

        }

        /// <summary>
        /// Gets or sets a formula for a range.
        /// </summary>
        public string Formula
        {
            get
            {
                if (IsName)
                {
                    if (_worksheet == null)
                    {
                        return _workbook._names[_address].NameFormula;
                    }
                    else
                    {
                        return _worksheet.Names[_address].NameFormula;
                    }
                }
                else
                {
                    return _worksheet.GetFormula(_fromRow, _fromCol);
                }
            }
            set
            {
                if (IsName)
                {
                    if (_worksheet == null)
                    {
                        _workbook._names[_address].NameFormula = value;
                    }
                    else
                    {
                        _worksheet.Names[_address].NameFormula = value;
                    }
                }
                else
                {
                    if (value == null || value.Trim() == "")
                    {
                        //Set the cells to null
                        Value = null;
                    }
                    else if (_fromRow == _toRow && _fromCol == _toCol)
                    {
                        Set_Formula(this, value, _fromRow, _fromCol);
                    }
                    else
                    {
                        Set_SharedFormula(this, value, this, false);
                        if (Addresses != null)
                        {
                            foreach (var address in Addresses)
                            {
                                Set_SharedFormula(this, value, address, false);
                            }
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
                IsRangeValid("FormulaR1C1");
                return _worksheet.GetFormulaR1C1(_fromRow, _fromCol);
            }
            set
            {
                IsRangeValid("FormulaR1C1");
                if (value.Length > 0 && value[0] == '=') value = value.Substring(1, value.Length - 1); // remove any starting equalsign.

                if (value == null || value.Trim() == "")
                {
                    //Set the cells to null
                    _worksheet.Cells[ExcelCellBase.TranslateFromR1C1(value, _fromRow, _fromCol)].Value = null;
                }
                else if (Addresses == null)
                {
                    Set_SharedFormula(this, ExcelCellBase.TranslateFromR1C1(value, _fromRow, _fromCol), this, false);
                }
                else
                {
                    Set_SharedFormula(this, ExcelCellBase.TranslateFromR1C1(value, _fromRow, _fromCol), new ExcelAddress(WorkSheet, FirstAddress), false);
                    foreach (var address in Addresses)
                    {
                        Set_SharedFormula(this, ExcelCellBase.TranslateFromR1C1(value, address.Start.Row, address.Start.Column), address, false);
                    }
                }
            }
        }
        /// <summary>
        /// Set the hyperlink property for a range of cells
        /// </summary>
        public Uri Hyperlink
        {
            get
            {
                IsRangeValid("formulaR1C1");
                return _worksheet._hyperLinks.GetValue(_fromRow, _fromCol);
            }
            set
            {
                _changePropMethod(this, _setHyperLinkDelegate, value);
            }
        }
        /// <summary>
        /// If the cells in the range are merged.
        /// </summary>
        public bool Merge
        {
            get
            {
                IsRangeValid("merging");
                for (int col = _fromCol; col <= _toCol; col++)
                {
                    for (int row = _fromRow; row <= _toRow; row++)
                    {
                        if (_worksheet.MergedCells[row, col] == null)
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            set
            {
                IsRangeValid("merging");
                _worksheet.MergedCells.Clear(this);
                if (value)
                {
                    _worksheet.MergedCells.Add(new ExcelAddressBase(FirstAddress), true);
                    if (Addresses != null)
                    {
                        foreach (var address in Addresses)
                        {
                            _worksheet.MergedCells.Clear(address); //Fixes issue 15482
                            _worksheet.MergedCells.Add(address, true);
                        }
                    }
                }
                else
                {
                    if (Addresses != null)
                    {
                        foreach (var address in Addresses)
                        {
                            _worksheet.MergedCells.Clear(address); ;
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
                IsRangeValid("autofilter");
                ExcelAddressBase address = _worksheet.AutoFilterAddress;
                if (address == null) return false;
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
                IsRangeValid("autofilter");
                if (_worksheet.AutoFilterAddress != null)
                {
                    var c = this.Collide(_worksheet.AutoFilterAddress);
                    if (value == false && (c == eAddressCollition.Partly || c == eAddressCollition.No))
                    {
                        throw (new InvalidOperationException("Can't remote Autofilter. Current autofilter does not match selected range."));
                    }
                }
                if (_worksheet.Names.ContainsKey("_xlnm._FilterDatabase"))
                {
                    _worksheet.Names.Remove("_xlnm._FilterDatabase");
                }
                if (value)
                {
                    _worksheet.AutoFilterAddress = this;
                    var result = _worksheet.Names.Add("_xlnm._FilterDatabase", this);
                    result.IsNameHidden = true;
                }
                else
                {
                    _worksheet.AutoFilterAddress = null;
                }
            }
        }
        /// <summary>
        /// If the value is in richtext format.
        /// </summary>
        public bool IsRichText
        {
            get
            {
                IsRangeValid("richtext");
                return _worksheet._flags.GetFlagValue(_fromRow, _fromCol, CellFlags.RichText);
            }
            set
            {
                _changePropMethod(this, _setIsRichTextDelegate, value);
            }
        }
        /// <summary>
        /// Is the range a part of an Arrayformula
        /// </summary>
        public bool IsArrayFormula
        {
            get
            {
                IsRangeValid("arrayformulas");
                return _worksheet._flags.GetFlagValue(_fromRow, _fromCol, CellFlags.ArrayFormula);
            }
        }
        protected ExcelRichTextCollection _rtc = null;
        /// <summary>
        /// Cell value is richtext formatted. 
        /// Richtext-property only apply to the left-top cell of the range.
        /// </summary>
        public ExcelRichTextCollection RichText
        {
            get
            {
                IsRangeValid("richtext");
                if (_rtc == null)
                {
                    _rtc = GetRichText(_fromRow, _fromCol);
                }
                return _rtc;
            }
        }

        private ExcelRichTextCollection GetRichText(int row, int col)
        {
            XmlDocument xml = new XmlDocument();
            var v = _worksheet.GetValueInner(row, col);
            var isRt = _worksheet._flags.GetFlagValue(row, col, CellFlags.RichText);
            if (v != null)
            {
                if (isRt)
                {
                    XmlHelper.LoadXmlSafe(xml, "<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" >" + v.ToString() + "</d:si>", Encoding.UTF8);
                }
                else
                {
                    xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ><d:r><d:t>" + OfficeOpenXml.Utils.ConvertUtil.ExcelEscapeString(v.ToString()) + "</d:t></d:r></d:si>");
                }
            }
            else
            {
                xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" />");
            }
            var rtc = new ExcelRichTextCollection(_worksheet.NameSpaceManager, xml.SelectSingleNode("d:si", _worksheet.NameSpaceManager), this);
            return rtc;
        }
        /// <summary>
        /// returns the comment object of the first cell in the range
        /// </summary>
        public ExcelComment Comment
        {
            get
            {
                IsRangeValid("comments");
                var i = -1;
                if (_worksheet.Comments.Count > 0)
                {
                    if (_worksheet._commentsStore.Exists(_fromRow, _fromCol, ref i))
                    {
                        return _worksheet._comments[i] as ExcelComment;
                    }
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
        public new string FullAddress
        {
            get
            {
                string fullAddress = GetFullAddress(_worksheet.Name, _address);
                if (Addresses != null)
                {
                    foreach (var a in Addresses)
                    {
                        fullAddress += "," + GetFullAddress(_worksheet.Name, a.Address); ;
                    }
                }
                return fullAddress;
            }
        }
        /// <summary>
        /// Address including sheetname
        /// </summary>
        public string FullAddressAbsolute
        {
            get
            {
                string wbwsRef = string.IsNullOrEmpty(base._wb) ? base._ws : "[" + base._wb.Replace("'", "''") + "]" + _ws;
                string fullAddress;
                if (Addresses == null)
                {
                    fullAddress = GetFullAddress(wbwsRef, GetAddress(_fromRow, _fromCol, _toRow, _toCol, true));
                }
                else
                {
                    fullAddress = "";
                    foreach (var a in Addresses)
                    {
                        if (fullAddress != "") fullAddress += ",";
                        if (a.Address == "#REF!")
                        {
                            fullAddress += GetFullAddress(wbwsRef, "#REF!");
                        }
                        else
                        {
                            fullAddress += GetFullAddress(wbwsRef, GetAddress(a.Start.Row, a.Start.Column, a.End.Row, a.End.Column, true));
                        }
                    }
                }
                return fullAddress;
            }
        }
        /// <summary>
        /// Address including sheetname
        /// </summary>
        internal string FullAddressAbsoluteNoFullRowCol
        {
            get
            {
                string wbwsRef = string.IsNullOrEmpty(base._wb) ? base._ws : "[" + base._wb.Replace("'", "''") + "]" + _ws;
                string fullAddress;
                if (Addresses == null)
                {
                    fullAddress = GetFullAddress(wbwsRef, GetAddress(_fromRow, _fromCol, _toRow, _toCol, true), false);
                }
                else
                {
                    fullAddress = "";
                    foreach (var a in Addresses)
                    {
                        if (fullAddress != "") fullAddress += ",";
                        fullAddress += GetFullAddress(wbwsRef, GetAddress(a.Start.Row, a.Start.Column, a.End.Row, a.End.Column, true), false); ;
                    }
                }
                return fullAddress;
            }
        }
        #endregion
        #region Private Methods
        /// <summary>
        /// Set the value without altering the richtext property
        /// </summary>
        /// <param name="value">the value</param>
        internal void SetValueRichText(object value)
        {
            if (_fromRow == 1 && _fromCol == 1 && _toRow == ExcelPackage.MaxRows && _toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
            {
                SetValue(value, 1, 1);
            }
            else
            {
                SetValue(value, _fromRow, _fromCol);
            }
        }

        private void SetValue(object value, int row, int col)
        {
            _worksheet.SetValue(row, col, value);
            // if (value is string) _worksheet._types.SetValue(row, col, "S"); else _worksheet._types.SetValue(row, col, "");
            _worksheet._formulas.SetValue(row, col, "");
        }
        internal void SetSharedFormulaID(int id)
        {
            for (int col = _fromCol; col <= _toCol; col++)
            {
                for (int row = _fromRow; row <= _toRow; row++)
                {
                    _worksheet._formulas.SetValue(row, col, id);
                }
            }
        }
        private void CheckAndSplitSharedFormula(ExcelAddressBase address)
        {
            for (int col = address._fromCol; col <= address._toCol; col++)
            {
                for (int row = address._fromRow; row <= address._toRow; row++)
                {
                    var f = _worksheet._formulas.GetValue(row, col);
                    if (f is int && (int)f >= 0)
                    {
                        SplitFormulas(address);
                        return;
                    }
                }
            }
        }

        private void SplitFormulas(ExcelAddressBase address)
        {
            List<int> formulas = new List<int>();
            for (int col = address._fromCol; col <= address._toCol; col++)
            {
                for (int row = address._fromRow; row <= address._toRow; row++)
                {
                    var f = _worksheet._formulas.GetValue(row, col);
                    if (f is int)
                    {
                        int id = (int)f;
                        if (id >= 0 && !formulas.Contains(id))
                        {
                            if (_worksheet._sharedFormulas[id].IsArray &&
                                    Collide(_worksheet.Cells[_worksheet._sharedFormulas[id].Address]) == eAddressCollition.Partly) // If the formula is an array formula and its on the inside the overwriting range throw an exception
                            {
                                throw (new InvalidOperationException("Can not overwrite a part of an array-formula"));
                            }
                            formulas.Add(id);
                        }
                    }
                }
            }

            foreach (int ix in formulas)
            {
                SplitFormula(address, ix);
            }

            ////Clear any formula references inside the refered range
            //_worksheet._formulas.Clear(address._fromRow, address._toRow, address._toRow - address._fromRow + 1, address._toCol - address.column + 1);
        }

        private void SplitFormula(ExcelAddressBase address, int ix)
        {
            var f = _worksheet._sharedFormulas[ix];
            var fRange = _worksheet.Cells[f.Address];
            var collide = address.Collide(fRange);

            //The formula is inside the currenct range, remove it
            if (collide == eAddressCollition.Equal || collide == eAddressCollition.Inside)
            {
                _worksheet._sharedFormulas.Remove(ix);
                return;
                //fRange.SetSharedFormulaID(int.MinValue); 
            }
            var firstCellCollide = address.Collide(new ExcelAddressBase(fRange._fromRow, fRange._fromCol, fRange._fromRow, fRange._fromCol));
            if (collide == eAddressCollition.Partly && (firstCellCollide == eAddressCollition.Inside || firstCellCollide == eAddressCollition.Equal)) //Do we need to split? Only if the functions first row is inside the new range.
            {
                //The formula partly collides with the current range
                bool fIsSet = false;
                string formulaR1C1 = fRange.FormulaR1C1;
                //Top Range
                if (fRange._fromRow < _fromRow)
                {
                    f.Address = ExcelCellBase.GetAddress(fRange._fromRow, fRange._fromCol, _fromRow - 1, fRange._toCol);
                    fIsSet = true;
                }
                //Left Range
                if (fRange._fromCol < address._fromCol)
                {
                    if (fIsSet)
                    {
                        f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
                        f.Index = _worksheet.GetMaxShareFunctionIndex(false);
                        f.StartCol = fRange._fromCol;
                        f.IsArray = false;
                        _worksheet._sharedFormulas.Add(f.Index, f);
                    }
                    else
                    {
                        fIsSet = true;
                    }
                    if (fRange._fromRow < address._fromRow)
                        f.StartRow = address._fromRow;
                    else
                    {
                        f.StartRow = fRange._fromRow;
                    }
                    if (fRange._toRow < address._toRow)
                    {
                        f.Address = ExcelCellBase.GetAddress(f.StartRow, f.StartCol,
                                fRange._toRow, address._fromCol - 1);
                    }
                    else
                    {
                        f.Address = ExcelCellBase.GetAddress(f.StartRow, f.StartCol,
                             address._toRow, address._fromCol - 1);
                    }
                    f.Formula = TranslateFromR1C1(formulaR1C1, f.StartRow, f.StartCol);
                    _worksheet.Cells[f.Address].SetSharedFormulaID(f.Index);
                }
                //Right Range
                if (fRange._toCol > address._toCol)
                {
                    if (fIsSet)
                    {
                        f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
                        f.Index = _worksheet.GetMaxShareFunctionIndex(false);
                        f.IsArray = false;
                        _worksheet._sharedFormulas.Add(f.Index, f);
                    }
                    else
                    {
                        fIsSet = true;
                    }
                    f.StartCol = address._toCol + 1;
                    if (address._fromRow < fRange._fromRow)
                        f.StartRow = fRange._fromRow;
                    else
                    {
                        f.StartRow = address._fromRow;
                    }

                    if (fRange._toRow < address._toRow)
                    {
                        f.Address = ExcelCellBase.GetAddress(f.StartRow, f.StartCol,
                                fRange._toRow, fRange._toCol);
                    }
                    else
                    {
                        f.Address = ExcelCellBase.GetAddress(f.StartRow, f.StartCol,
                                address._toRow, fRange._toCol);
                    }
                    f.Formula = TranslateFromR1C1(formulaR1C1, f.StartRow, f.StartCol);
                    _worksheet.Cells[f.Address].SetSharedFormulaID(f.Index);
                }
                //Bottom Range
                if (fRange._toRow > address._toRow)
                {
                    if (fIsSet)
                    {
                        f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
                        f.Index = _worksheet.GetMaxShareFunctionIndex(false);
                        f.IsArray = false;
                        _worksheet._sharedFormulas.Add(f.Index, f);
                    }

                    f.StartCol = fRange._fromCol;
                    f.StartRow = _toRow + 1;

                    f.Formula = TranslateFromR1C1(formulaR1C1, f.StartRow, f.StartCol);

                    f.Address = ExcelCellBase.GetAddress(f.StartRow, f.StartCol,
                            fRange._toRow, fRange._toCol);
                    _worksheet.Cells[f.Address].SetSharedFormulaID(f.Index);

                }
            }
        }
        private object ConvertData(ExcelTextFormat Format, string v, int col, bool isText)
        {
            if (isText && (Format.DataTypes == null || Format.DataTypes.Length < col)) return string.IsNullOrEmpty(v) ? null : v;

            double d;
            DateTime dt;
            if (Format.DataTypes == null || Format.DataTypes.Length <= col || Format.DataTypes[col] == eDataTypes.Unknown)
            {
                string v2 = v.EndsWith("%") ? v.Substring(0, v.Length - 1) : v;
                if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out d))
                {
                    if (v2 == v)
                    {
                        return d;
                    }
                    else
                    {
                        return d / 100;
                    }
                }
                if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out dt))
                {
                    return dt;
                }
                else
                {
                    return string.IsNullOrEmpty(v) ? null : v; ;
                }
            }
            else
            {
                switch (Format.DataTypes[col])
                {
                    case eDataTypes.Number:
                        if (double.TryParse(v, NumberStyles.Any, Format.Culture, out d))
                        {
                            return d;
                        }
                        else
                        {
                            return v;
                        }
                    case eDataTypes.DateTime:
                        if (DateTime.TryParse(v, Format.Culture, DateTimeStyles.None, out dt))
                        {
                            return dt;
                        }
                        else
                        {
                            return v;
                        }
                    case eDataTypes.Percent:
                        string v2 = v.EndsWith("%") ? v.Substring(0, v.Length - 1) : v;
                        if (double.TryParse(v2, NumberStyles.Any, Format.Culture, out d))
                        {
                            return d / 100;
                        }
                        else
                        {
                            return v;
                        }
                    case eDataTypes.String:
                        return v;
                    default:
                        return string.IsNullOrEmpty(v) ? null : v;

                }
            }
        }
        #endregion
        #region Public Methods
        #region ConditionalFormatting
        /// <summary>
        /// Conditional Formatting for this range.
        /// </summary>
        public IRangeConditionalFormatting ConditionalFormatting
        {
            get
            {
                return new RangeConditionalFormatting(_worksheet, new ExcelAddress(Address));
            }
        }
        #endregion
        #region DataValidation
        /// <summary>
        /// Data validation for this range.
        /// </summary>
        public IRangeDataValidation DataValidation
        {
            get
            {
                return new RangeDataValidation(_worksheet, Address);
            }
        }
        #endregion
        #region LoadFromDataReader
        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to loadfrom</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableName">The name of the table</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders, string TableName, TableStyles TableStyle = TableStyles.None)
        {
            var r = LoadFromDataReader(Reader, PrintHeaders);

            int rows = r.Rows - 1;
            if (rows >= 0 && r.Columns > 0)
            {
                var tbl = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + (rows <= 0 ? 1 : rows), _fromCol + r.Columns - 1), TableName);
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }
            return r;
        }

        /// <summary>
        /// Load the data from the datareader starting from the top left cell of the range
        /// </summary>
        /// <param name="Reader">The datareader to load from</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataReader(IDataReader Reader, bool PrintHeaders)
        {
            if (Reader == null)
            {
                throw (new ArgumentNullException("Reader", "Reader can't be null"));
            }
            int fieldCount = Reader.FieldCount;

            int col = _fromCol, row = _fromRow;
            if (PrintHeaders)
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    // If no caption is set, the ColumnName property is called implicitly.
                    _worksheet.SetValueInner(row, col++, Reader.GetName(i));
                }
                row++;
                col = _fromCol;
            }
            while (Reader.Read())
            {
                for (int i = 0; i < fieldCount; i++)
                {
                    _worksheet.SetValueInner(row, col++, Reader.GetValue(i));
                }
                row++;
                col = _fromCol;
            }
            return _worksheet.Cells[_fromRow, _fromCol, row - 1, _fromCol + fieldCount - 1];
        }
        #endregion

        #region LoadFromDataTable
        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">Print the column caption property (if set) or the columnname property if not, on first row</param>
        /// <param name="TableStyle">The table style to apply to the data</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders, TableStyles TableStyle)
        {
            var r = LoadFromDataTable(Table, PrintHeaders);

            int rows = (Table.Rows.Count == 0 ? 1 : Table.Rows.Count) + (PrintHeaders ? 1 : 0);
            if (rows >= 0 && Table.Columns.Count > 0)
            {
                var tbl = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + rows - 1, _fromCol + Table.Columns.Count - 1), Table.TableName);
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }
            return r;
        }
        /// <summary>
        /// Load the data from the datatable starting from the top left cell of the range
        /// </summary>
        /// <param name="Table">The datatable to load</param>
        /// <param name="PrintHeaders">Print the caption property (if set) or the columnname property if not, on first row</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromDataTable(DataTable Table, bool PrintHeaders)
        {
            if (Table == null)
            {
                throw (new ArgumentNullException("Table can't be null"));
            }

            if (Table.Rows.Count == 0 && PrintHeaders == false)
            {
                return null;
            }

            var rowArray = new List<object[]>();
            if (PrintHeaders)
            {
                rowArray.Add(Table.Columns.Cast<DataColumn>().Select((dc) => { return dc.Caption; }).ToArray());
            }
            foreach (DataRow dr in Table.Rows)
            {
                rowArray.Add(dr.ItemArray);
            }
            _worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + Table.Columns.Count - 1,
                (List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object value) =>
                {
                    rowIx -= _fromRow;
                    columnIx -= _fromCol;

                    var val = ((List<object[]>)value)[rowIx][columnIx];
                    if (val != null && val != DBNull.Value && !string.IsNullOrEmpty(val.ToString()))
                    {
                        list[index] = new ExcelCoreValue { _value = val, _styleId = list[index]._styleId };
                    }
                }, rowArray);

            return _worksheet.Cells[_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + Table.Columns.Count - 1];
        }
        #endregion

        #region LoadFromArrays
        /// <summary>
        /// Loads data from the collection of arrays of objects into the range, starting from
        /// the top-left cell.
        /// </summary>
        /// <param name="Data">The data.</param>
        public ExcelRangeBase LoadFromArrays(IEnumerable<object[]> Data)
        {
            //thanx to Abdullin for the code contribution
            if (Data == null) throw new ArgumentNullException("data");

            var rowArray = new List<object[]>();
            var maxColumn = 0;
            foreach (object[] item in Data)
            {
                rowArray.Add(item);
                if (maxColumn < item.Length) maxColumn = item.Length;
            }
            if (rowArray.Count == 0) return null; //Issue #57
            _worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + maxColumn - 1,
                (List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object value) =>
                {
                    rowIx -= _fromRow;
                    columnIx -= _fromCol;

                    var values = ((List<object[]>)value);
                    if (values.Count <= rowIx) return;
                    var item = values[rowIx];
                    if (item.Length <= columnIx) return;

                    var val = item[columnIx];
                    if (val != null && val != DBNull.Value && !string.IsNullOrEmpty(val.ToString()))
                    {
                        list[index] = new ExcelCoreValue { _value = val, _styleId = list[index]._styleId };
                    }
                }, rowArray);

            return _worksheet.Cells[_fromRow, _fromCol, _fromRow + rowArray.Count - 1, _fromCol + maxColumn - 1];
        }
        #endregion
        #region LoadFromCollection
        /// <summary>
        /// Load a collection into a the worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection)
        {
            return LoadFromCollection<T>(Collection, false, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, null);
        }
        /// <summary>
        /// Load a collection of T into the worksheet starting from the top left row of the range.
        /// Default option will load all public instance properties of T
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders)
        {
            return LoadFromCollection<T>(Collection, PrintHeaders, TableStyles.None, BindingFlags.Public | BindingFlags.Instance, null);
        }
        /// <summary>
        /// Load a collection of T into the worksheet starting from the top left row of the range.
        /// Default option will load all public instance properties of T
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles TableStyle)
        {
            return LoadFromCollection<T>(Collection, PrintHeaders, TableStyle, BindingFlags.Public | BindingFlags.Instance, null);
        }
        /// <summary>
        /// Load a collection into the worksheet starting from the top left row of the range.
        /// </summary>
        /// <typeparam name="T">The datatype in the collection</typeparam>
        /// <param name="Collection">The collection to load</param>
        /// <param name="PrintHeaders">Print the property names on the first row. Any underscore in the property name will be converted to a space. If the property is decorated with a <see cref="DisplayNameAttribute"/> or a <see cref="DescriptionAttribute"/> that attribute will be used instead of the reflected member name.</param>
        /// <param name="TableStyle">Will create a table with this style. If set to TableStyles.None no table will be created</param>
        /// <param name="memberFlags">Property flags to use</param>
        /// <param name="Members">The properties to output. Must be of type T</param>
        /// <returns>The filled range</returns>
        public ExcelRangeBase LoadFromCollection<T>(IEnumerable<T> Collection, bool PrintHeaders, TableStyles TableStyle, BindingFlags memberFlags, MemberInfo[] Members)
        {
            var type = typeof(T);
            bool isSameType = true;
            if (Members == null)
            {
                Members = type.GetProperties(memberFlags);
            }
            else
            {
                if (Members.Length == 0)   //Fixes issue 15555
                {
                    throw (new ArgumentException("Parameter Members must have at least one property. Length is zero"));
                }
                foreach (var t in Members)
                {
                    if (t.DeclaringType != null && t.DeclaringType != type)
                    {
                        isSameType = false;
                    }
                    //Fixing inverted check for IsSubclassOf / Pullrequest from tomdam
                    if (t.DeclaringType != null && t.DeclaringType != type && !TypeCompat.IsSubclassOf(type, t.DeclaringType) && !TypeCompat.IsSubclassOf(t.DeclaringType, type))
                    {
                        throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T)");
                    }
                }
            }

            // create buffer
            object[,] values = new object[(PrintHeaders ? Collection.Count() + 1 : Collection.Count()), Members.Count()];

            int col = 0, row = 0;
            if (Members.Length > 0 && PrintHeaders)
            {
                foreach (var t in Members)
                {
                    var descriptionAttribute = t.GetCustomAttributes(typeof(DescriptionAttribute), false).FirstOrDefault() as DescriptionAttribute;
                    var header = string.Empty;
                    if (descriptionAttribute != null)
                    {
                        header = descriptionAttribute.Description;
                    }
                    else
                    {
                        var displayNameAttribute =
                            t.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault() as
                            DisplayNameAttribute;
                        if (displayNameAttribute != null)
                        {
                            header = displayNameAttribute.DisplayName;
                        }
                        else
                        {
                            header = t.Name.Replace('_', ' ');
                        }
                    }
                    //_worksheet.SetValueInner(row, col++, header);
                    values[row, col++] = header;
                }
                row++;
            }

            if (!Collection.Any() && (Members.Length == 0 || PrintHeaders == false))
            {
                return null;
            }

            foreach (var item in Collection)
            {
                col = 0;
                if (item is string || item is decimal || item is DateTime || TypeCompat.IsPrimitive(item))
                {
                    values[row, col++] = item;
                }
                else
                {
                    foreach (var t in Members)
                    {
                        if (isSameType == false && item.GetType().GetMember(t.Name, memberFlags).Length == 0)
                        {
                            col++;
                            continue; //Check if the property exists if and inherited class is used
                        }
                        else if (t is PropertyInfo)
                        {
                            values[row, col++] = ((PropertyInfo)t).GetValue(item, null);
                        }
                        else if (t is FieldInfo)
                        {
                            values[row, col++] = ((FieldInfo)t).GetValue(item);
                        }
                        else if (t is MethodInfo)
                        {
                            values[row, col++] = ((MethodInfo)t).Invoke(item, null);
                        }
                    }
                }
                row++;
            }

            _worksheet.SetRangeValueInner(_fromRow, _fromCol, _fromRow + row - 1, _fromCol + col - 1, values);

            //Must have at least 1 row, if header is showen
            if (row == 1 && PrintHeaders)
            {
                row++;
            }

            var r = _worksheet.Cells[_fromRow, _fromCol, _fromRow + row - 1, _fromCol + col - 1];

            if (TableStyle != TableStyles.None)
            {
                var tbl = _worksheet.Tables.Add(r, "");
                tbl.ShowHeader = PrintHeaders;
                tbl.TableStyle = TableStyle;
            }
            return r;
        }
        #endregion
        #region LoadFromText
        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// Default settings is Comma separation
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <returns>The range containing the data</returns>
        public ExcelRangeBase LoadFromText(string Text)
        {
            return LoadFromText(Text, new ExcelTextFormat());
        }
        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns>The range containing the data</returns>
        public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format)
        {
            if (string.IsNullOrEmpty(Text))
            {
                var r = _worksheet.Cells[_fromRow, _fromCol];
                r.Value = "";
                return r;
            }

            if (Format == null) Format = new ExcelTextFormat();


            string[] lines;
            if (Format.TextQualifier == 0)
            {
                lines = Regex.Split(Text, Format.EOL);
            }
            else
            {
                lines = GetLines(Text, Format);
            }
            //string splitRegex = String.Format("{0}(?=(?:[^{1}]*{1}[^{1}]*{1})*[^{1}]*$)", Format.EOL, Format.TextQualifier);
            //lines = Regex.Split(Text, splitRegex);

            int row = 0;
            int col = 0;
            int maxCol = col;
            int lineNo = 1;
            var values = new List<object>[lines.Length];
            foreach (string line in lines)
            {
                var items = new List<object>();
                values[row] = items;

                if (lineNo > Format.SkipLinesBeginning && lineNo <= lines.Length - Format.SkipLinesEnd)
                {
                    col = 0;
                    string v = "";
                    bool isText = false, isQualifier = false;
                    int QCount = 0;
                    int lineQCount = 0;
                    foreach (char c in line)
                    {
                        if (Format.TextQualifier != 0 && c == Format.TextQualifier)
                        {
                            if (!isText && v != "")
                            {
                                throw (new Exception(string.Format("Invalid Text Qualifier in line : {0}", line)));
                            }
                            isQualifier = !isQualifier;
                            QCount += 1;
                            lineQCount++;
                            isText = true;
                        }
                        else
                        {
                            if (QCount > 1 && !string.IsNullOrEmpty(v))
                            {
                                v += new string(Format.TextQualifier, QCount / 2);
                            }
                            else if (QCount > 2 && string.IsNullOrEmpty(v))
                            {
                                v += new string(Format.TextQualifier, (QCount - 1) / 2);
                            }

                            if (isQualifier)
                            {
                                v += c;
                            }
                            else
                            {
                                if (c == Format.Delimiter)
                                {
                                    items.Add(ConvertData(Format, v, col, isText));
                                    v = "";
                                    isText = false;
                                    col++;
                                }
                                else
                                {
                                    if (QCount % 2 == 1)
                                    {
                                        throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));
                                    }
                                    v += c;
                                }
                            }
                            QCount = 0;
                        }
                    }
                    if (QCount > 1 && (v != "" && QCount == 2))
                    {
                        v += new string(Format.TextQualifier, QCount / 2);
                    }
                    if (lineQCount % 2 == 1)
                        throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));

                    //_worksheet.SetValueInner(row, col, ConvertData(Format, v, col - _fromCol, isText));
                    items.Add(ConvertData(Format, v, col, isText));
                    if (col > maxCol) maxCol = col;
                    row++;
                }
                lineNo++;
            }
            // flush
            _worksheet._values.SetRangeValueSpecial(_fromRow, _fromCol, _fromRow + values.Length - 1, _fromCol + maxCol,
                (List<ExcelCoreValue> list, int index, int rowIx, int columnIx, object value) =>
                {
                    rowIx -= _fromRow;
                    columnIx -= _fromCol;
                    var item = values[rowIx];
                    if (item == null || item.Count <= columnIx) return;

                    list[index] = new ExcelCoreValue { _value = item[columnIx], _styleId = list[index]._styleId };
                }, values);

            return _worksheet.Cells[_fromRow, _fromCol, _fromRow + row - 1, _fromCol + maxCol];
        }

        private string[] GetLines(string text, ExcelTextFormat Format)
        {
            if (Format.EOL == null || Format.EOL.Length == 0) return new string[] { text };
            var eol = Format.EOL;
            var list = new List<string>();
            var inTQ = false;
            var prevLineStart = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == Format.TextQualifier)
                {
                    inTQ = !inTQ;
                }
                else if (!inTQ)
                {
                    if (IsEOL(text, i, eol))
                    {
                        list.Add(text.Substring(prevLineStart, i - prevLineStart));
                        i += eol.Length - 1;
                        prevLineStart = i + 1;
                    }
                }
            }

            if (inTQ)
            {
                throw (new ArgumentException(string.Format("Text delimiter is not closed in line : {0}", list.Count)));
            }

            if (prevLineStart >= Format.EOL.Length && IsEOL(text, prevLineStart - Format.EOL.Length, Format.EOL))
            {
                //list.Add(text.Substring(prevLineStart- Format.EOL.Length, Format.EOL.Length));
                list.Add("");
            }
            else
            {
                list.Add(text.Substring(prevLineStart));
            }
            return list.ToArray();
        }
        private bool IsEOL(string text, int ix, string eol)
        {
            for (int i = 0; i < eol.Length; i++)
            {
                if (text[ix + i] != eol[i])
                    return false;
            }
            return ix + eol.Length <= text.Length;
        }

        /// <summary>
        /// Loads a CSV text into a range starting from the top left cell.
        /// </summary>
        /// <param name="Text">The Text</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(string Text, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader)
        {
            var r = LoadFromText(Text, Format);

            var tbl = _worksheet.Tables.Add(r, "");
            tbl.ShowHeader = FirstRowIsHeader;
            tbl.TableStyle = TableStyle;

            return r;
        }
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile)
        {
            return LoadFromText(File.ReadAllText(TextFile.FullName, Encoding.ASCII));
        }
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format)
        {
            return LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format);
        }
        /// <summary>
        /// Loads a CSV file into a range starting from the top left cell.
        /// </summary>
        /// <param name="TextFile">The Textfile</param>
        /// <param name="Format">Information how to load the text</param>
        /// <param name="TableStyle">Create a table with this style</param>
        /// <param name="FirstRowIsHeader">Use the first row as header</param>
        /// <returns></returns>
        public ExcelRangeBase LoadFromText(FileInfo TextFile, ExcelTextFormat Format, TableStyles TableStyle, bool FirstRowIsHeader)
        {
            return LoadFromText(File.ReadAllText(TextFile.FullName, Format.Encoding), Format, TableStyle, FirstRowIsHeader);
        }
        #endregion
        #region GetValue

        /// <summary>
        ///     Convert cell value to desired type, including nullable structs.
        ///     When converting blank string to nullable struct (e.g. ' ' to int?) null is returned.
        ///     When attempted conversion fails exception is passed through.
        /// </summary>
        /// <typeparam name="T">
        ///     The type to convert to.
        /// </typeparam>
        /// <returns>
        ///     The <see cref="Value"/> converted to <typeparamref name="T"/>.
        /// </returns>
        /// <remarks>
        ///     If  <see cref="Value"/> is string, parsing is performed for output types of DateTime and TimeSpan, which if fails throws <see cref="FormatException"/>.
        ///     Another special case for output types of DateTime and TimeSpan is when input is double, in which case <see cref="DateTime.FromOADate"/>
        ///     is used for conversion. This special case does not work through other types convertible to double (e.g. integer or string with number).
        ///     In all other cases 'direct' conversion <see cref="Convert.ChangeType(object, Type)"/> is performed.
        /// </remarks>
        /// <exception cref="FormatException">
        ///      <see cref="Value"/> is string and its format is invalid for conversion (parsing fails)
        /// </exception>
        /// <exception cref="InvalidCastException">
        ///      <see cref="Value"/> is not string and direct conversion fails
        /// </exception>
        public T GetValue<T>()
        {
            return ConvertUtil.GetTypedCellValue<T>(Value);
        }
        #endregion
        /// <summary>
        /// Get a range with an offset from the top left cell.
        /// The new range has the same dimensions as the current range
        /// </summary>
        /// <param name="RowOffset">Row Offset</param>
        /// <param name="ColumnOffset">Column Offset</param>
        /// <returns></returns>
        public ExcelRangeBase Offset(int RowOffset, int ColumnOffset)
        {
            if (_fromRow + RowOffset < 1 || _fromCol + ColumnOffset < 1 || _fromRow + RowOffset > ExcelPackage.MaxRows || _fromCol + ColumnOffset > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentOutOfRangeException("Offset value out of range"));
            }
            string address = GetAddress(_fromRow + RowOffset, _fromCol + ColumnOffset, _toRow + RowOffset, _toCol + ColumnOffset);
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
            if (_fromRow + RowOffset < 1 || _fromCol + ColumnOffset < 1 || _fromRow + RowOffset > ExcelPackage.MaxRows || _fromCol + ColumnOffset > ExcelPackage.MaxColumns ||
                 _fromRow + RowOffset + NumberOfRows < 1 || _fromCol + ColumnOffset + NumberOfColumns < 1 || _fromRow + RowOffset + NumberOfRows > ExcelPackage.MaxRows || _fromCol + ColumnOffset + NumberOfColumns > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentOutOfRangeException("Offset value out of range"));
            }
            string address = GetAddress(_fromRow + RowOffset, _fromCol + ColumnOffset, _fromRow + RowOffset + NumberOfRows, _fromCol + ColumnOffset + NumberOfColumns);
            return new ExcelRangeBase(_worksheet, address);
        }
        /// <summary>
        /// Adds a new comment for the range.
        /// If this range contains more than one cell, the top left comment is returned by the method.
        /// </summary>
        /// <param name="Text"></param>
        /// <param name="Author"></param>
        /// <returns>A reference comment of the top left cell</returns>
        public ExcelComment AddComment(string Text, string Author)
        {
            if (string.IsNullOrEmpty(Author))
            {
#if Core
                Author = System.Security.Claims.ClaimsPrincipal.Current.Identity.Name;
#else
                Author = Thread.CurrentPrincipal.Identity.Name;
#endif
            }
            //Check if any comments exists in the range and throw an exception
            _changePropMethod(this, _setExistsCommentDelegate, null);
            //Create the comments
            _changePropMethod(this, _setCommentDelegate, new string[] { Text, Author });

            return _worksheet.Comments[new ExcelCellAddress(_fromRow, _fromCol)];
        }

        /// <summary>
        /// Copies the range of cells to an other range
        /// </summary>
        /// <param name="Destination">The start cell where the range will be copied.</param>
        public void Copy(ExcelRangeBase Destination)
        {
            Copy(Destination, null);
        }

        /// <summary>
        /// Copies the range of cells to an other range
        /// </summary>
        /// <param name="Destination">The start cell where the range will be copied.</param>
        /// <param name="excelRangeCopyOptionFlags">Cell parts that will not be copied. If Formulas are specified, the formulas will NOT be copied.</param>
        public void Copy(ExcelRangeBase Destination, ExcelRangeCopyOptionFlags? excelRangeCopyOptionFlags)
        {
            bool sameWorkbook = Destination._worksheet.Workbook == _worksheet.Workbook;
            ExcelStyles sourceStyles = _worksheet.Workbook.Styles,
                        styles = Destination._worksheet.Workbook.Styles;
            Dictionary<int, int> styleCashe = new Dictionary<int, int>();

            //Clear all existing cells; 
            int toRow = _toRow - _fromRow + 1,
                toCol = _toCol - _fromCol + 1;

            int i = 0;
            object o = null;
            byte flag = 0;
            Uri hl = null;

            var excludeFormulas = excelRangeCopyOptionFlags.HasValue && (excelRangeCopyOptionFlags.Value & ExcelRangeCopyOptionFlags.ExcludeFormulas) == ExcelRangeCopyOptionFlags.ExcludeFormulas;
            var cse = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);

            var copiedValue = new List<CopiedCell>();
            while (cse.Next())
            {
                var row = cse.Row;
                var col = cse.Column;       //Issue 15070
                var cell = new CopiedCell
                {
                    Row = Destination._fromRow + (row - _fromRow),
                    Column = Destination._fromCol + (col - _fromCol),
                    Value = cse.Value._value
                };

                if (!excludeFormulas && _worksheet._formulas.Exists(row, col, ref o))
                {
                    if (o is int)
                    {
                        cell.Formula = _worksheet.GetFormula(cse.Row, cse.Column);
                        if (_worksheet._flags.GetFlagValue(cse.Row, cse.Column, CellFlags.ArrayFormula))
                        {
                            Destination._worksheet._flags.SetFlagValue(cse.Row, cse.Column, true, CellFlags.ArrayFormula);
                        }
                    }
                    else
                    {
                        cell.Formula = o;
                    }
                }
                if (_worksheet.ExistsStyleInner(row, col, ref i))
                {
                    if (sameWorkbook)
                    {
                        cell.StyleID = i;
                    }
                    else
                    {
                        if (styleCashe.ContainsKey(i))
                        {
                            i = styleCashe[i];
                        }
                        else
                        {
                            var oldStyleID = i;
                            i = styles.CloneStyle(sourceStyles, i);
                            styleCashe.Add(oldStyleID, i);
                        }
                        cell.StyleID = i;
                    }
                }

                if (_worksheet._hyperLinks.Exists(row, col, ref hl))
                {
                    cell.HyperLink = hl;
                }

                // Will just be null if no comment exists.
                cell.Comment = _worksheet.Cells[cse.Row, cse.Column].Comment;

                if (_worksheet._flags.Exists(row, col, ref flag))
                {
                    cell.Flag = flag;
                }
                copiedValue.Add(cell);
            }

            //Copy styles with no cell value
            var cses = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
            while (cses.Next())
            {
                if (!_worksheet.ExistsValueInner(cses.Row, cses.Column))
                {
                    var row = Destination._fromRow + (cses.Row - _fromRow);
                    var col = Destination._fromCol + (cses.Column - _fromCol);
                    var cell = new CopiedCell
                    {
                        Row = row,
                        Column = col,
                        Value = null
                    };

                    i = cses.Value._styleId;
                    if (sameWorkbook)
                    {
                        cell.StyleID = i;
                    }
                    else
                    {
                        if (styleCashe.ContainsKey(i))
                        {
                            i = styleCashe[i];
                        }
                        else
                        {
                            var oldStyleID = i;
                            i = styles.CloneStyle(sourceStyles, i);
                            styleCashe.Add(oldStyleID, i);
                        }
                        cell.StyleID = i;
                    }
                    copiedValue.Add(cell);
                }
            }
            var copiedMergedCells = new Dictionary<int, ExcelAddress>();
            //Merged cells
            var csem = new CellsStoreEnumerator<int>(_worksheet.MergedCells._cells, _fromRow, _fromCol, _toRow, _toCol);
            while (csem.Next())
            {
                if (!copiedMergedCells.ContainsKey(csem.Value))
                {
                    var adr = new ExcelAddress(_worksheet.Name, _worksheet.MergedCells.List[csem.Value]);
                    var collideResult = Collide(adr);
                    if (collideResult == eAddressCollition.Inside || collideResult == eAddressCollition.Equal)
                    {
                        copiedMergedCells.Add(csem.Value, new ExcelAddress(
                            Destination._fromRow + (adr.Start.Row - _fromRow),
                            Destination._fromCol + (adr.Start.Column - _fromCol),
                            Destination._fromRow + (adr.End.Row - _fromRow),
                            Destination._fromCol + (adr.End.Column - _fromCol)));
                    }
                    else
                    {
                        //Partial merge of the address ignore.
                        copiedMergedCells.Add(csem.Value, null);
                    }
                }
            }

            Destination._worksheet.MergedCells.Clear(new ExcelAddressBase(Destination._fromRow, Destination._fromCol, Destination._fromRow + toRow - 1, Destination._fromCol + toCol - 1));

            Destination._worksheet._values.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._formulas.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._hyperLinks.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._flags.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._commentsStore.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);

            foreach (var cell in copiedValue)
            {
                Destination._worksheet.SetValueInner(cell.Row, cell.Column, cell.Value);

                if (cell.StyleID != null)
                {
                    Destination._worksheet.SetStyleInner(cell.Row, cell.Column, cell.StyleID.Value);
                }

                if (cell.Formula != null)
                {
                    cell.Formula = UpdateFormulaReferences(cell.Formula.ToString(), Destination._fromRow - _fromRow, Destination._fromCol - _fromCol, 0, 0, Destination.WorkSheet, Destination.WorkSheet, true);
                    Destination._worksheet._formulas.SetValue(cell.Row, cell.Column, cell.Formula);
                }
                if (cell.HyperLink != null)
                {
                    Destination._worksheet._hyperLinks.SetValue(cell.Row, cell.Column, cell.HyperLink);
                }

                if (cell.Comment != null)
                {
                    Destination.Worksheet.Cells[cell.Row, cell.Column].AddComment(cell.Comment.Text, cell.Comment.Author);
                }
                if (cell.Flag != 0)
                {
                    Destination._worksheet._flags.SetValue(cell.Row, cell.Column, cell.Flag);
                }
            }

            //Add merged cells
            foreach (var m in copiedMergedCells.Values)
            {
                if (m != null)
                {
                    Destination._worksheet.MergedCells.Add(m, true);
                }
            }

            //Check that the range is not larger than the dimensions of the worksheet. 
            //If so set the copied range to the worksheet dimensions to avoid copying empty cells.
            ExcelAddressBase range;

            if (Worksheet.Dimension == null)
            {
                range = this;
            }
            else
            {
                var collideStatus = Collide(Worksheet.Dimension);
                if (collideStatus != eAddressCollition.Equal || collideStatus != eAddressCollition.Inside)
                {
                    range = Worksheet.Dimension;
                }
                else
                {
                    range = this;
                }
            }

            if (_fromCol == 1 && _toCol == ExcelPackage.MaxColumns)
            {
                for (int r = 0; r < range.Rows; r++)
                {
                    var destinationRow = Destination.Worksheet.Row(Destination.Start.Row + r);
                    destinationRow.OutlineLevel = Worksheet.Row(range._fromRow + r).OutlineLevel;
                }
            }
            if (_fromRow == 1 && _toRow == ExcelPackage.MaxRows)
            {
                for (int c = 0; c < range.Columns; c++)
                {
                    var destinationCol = Destination.Worksheet.Column(Destination.Start.Column + c);
                    destinationCol.OutlineLevel = Worksheet.Column(range._fromCol + c).OutlineLevel;
                }
            }
        }

        /// <summary>
        /// Clear all cells
        /// </summary>
        public void Clear()
        {
            Delete(this, false);
        }
        /// <summary>
        /// Creates an array-formula.
        /// </summary>
        /// <param name="ArrayFormula">The formula</param>
        public void CreateArrayFormula(string ArrayFormula)
        {
            if (Addresses != null)
            {
                throw (new Exception("An Arrayformula can not have more than one address"));
            }
            Set_SharedFormula(this, ArrayFormula, this, true);
        }
        //private void Clear(ExcelAddressBase Range)
        //{
        //    Clear(Range, true);
        //}
        internal void Delete(ExcelAddressBase Range, bool shift)
        {
            //DeleteCheckMergedCells(Range);
            _worksheet.MergedCells.Clear(Range);
            //First find the start cell
            int fromRow, fromCol;
            var d = Worksheet.Dimension;
            if (d != null && Range._fromRow <= d._fromRow && Range._toRow >= d._toRow) //EntireRow?
            {
                fromRow = 0;
            }
            else
            {
                fromRow = Range._fromRow;
            }
            if (d != null && Range._fromCol <= d._fromCol && Range._toCol >= d._toCol) //EntireRow?
            {
                fromCol = 0;
            }
            else
            {
                fromCol = Range._fromCol;
            }

            var rows = Range._toRow - fromRow + 1;
            var cols = Range._toCol - fromCol + 1;

            _worksheet._values.Delete(fromRow, fromCol, rows, cols, shift);
            //_worksheet._types.Delete(fromRow, fromCol, rows, cols, shift);
            //_worksheet._styles.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._formulas.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._hyperLinks.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._flags.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._commentsStore.Delete(fromRow, fromCol, rows, cols, shift);

            //Clear multi addresses as well
            if (Addresses != null)
            {
                foreach (var sub in Addresses)
                {
                    Delete(sub, shift);
                }
            }
        }

        private void DeleteCheckMergedCells(ExcelAddressBase Range)
        {
            var removeItems = new List<string>();
            foreach (var addr in Worksheet.MergedCells)
            {
                var addrCol = Range.Collide(new ExcelAddress(Range.WorkSheet, addr));
                if (addrCol != eAddressCollition.No)
                {
                    if (addrCol == eAddressCollition.Inside)
                    {
                        removeItems.Add(addr);
                    }
                    else
                    {
                        throw (new InvalidOperationException("Can't remove/overwrite a part of cells that are merged"));
                    }
                }
            }
            foreach (var item in removeItems)
            {
                Worksheet.MergedCells.Remove(item);
            }
        }
        #endregion
        #region IDisposable Members

        public void Dispose()
        {
            //_worksheet = null;            
        }

        #endregion
        #region "Enumerator"
        CellsStoreEnumerator<ExcelCoreValue> cellEnum;
        public IEnumerator<ExcelRangeBase> GetEnumerator()
        {
            Reset();
            return this;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            Reset();
            return this;
        }

        /// <summary>
        /// The current range when enumerating
        /// </summary>
        public ExcelRangeBase Current
        {
            get
            {
                return new ExcelRangeBase(_worksheet, ExcelAddressBase.GetAddress(cellEnum.Row, cellEnum.Column));
            }
        }

        /// <summary>
        /// The current range when enumerating
        /// </summary>
        object IEnumerator.Current
        {
            get
            {
                return ((object)(new ExcelRangeBase(_worksheet, ExcelAddressBase.GetAddress(cellEnum.Row, cellEnum.Column))));
            }
        }

        //public object FormatedText { get; private set; }

        int _enumAddressIx = -1;
        public bool MoveNext()
        {
            if (cellEnum.Next())
            {
                return true;
            }
            else if (_addresses != null)
            {
                _enumAddressIx++;
                if (_enumAddressIx < _addresses.Count)
                {
                    cellEnum = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values,
                        _addresses[_enumAddressIx]._fromRow,
                        _addresses[_enumAddressIx]._fromCol,
                        _addresses[_enumAddressIx]._toRow,
                        _addresses[_enumAddressIx]._toCol);
                    return MoveNext();
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

        public void Reset()
        {
            _enumAddressIx = -1;
            cellEnum = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
        }
        #endregion
        private struct SortItem<T>
        {
            internal int Row { get; set; }
            internal T[] Items { get; set; }
        }
        private class Comp : IComparer<SortItem<ExcelCoreValue>>
        {
            public int[] columns;
            public bool[] descending;
            public CultureInfo cultureInfo = CultureInfo.CurrentCulture;
            public CompareOptions compareOptions = CompareOptions.None;
            public int Compare(SortItem<ExcelCoreValue> x, SortItem<ExcelCoreValue> y)
            {
                var ret = 0;
                for (int i = 0; i < columns.Length; i++)
                {
                    var x1 = x.Items[columns[i]]._value;
                    var y1 = y.Items[columns[i]]._value;
                    var isNumX = ConvertUtil.IsNumeric(x1);
                    var isNumY = ConvertUtil.IsNumeric(y1);
                    if (isNumX && isNumY)   //Numeric Compare
                    {
                        var d1 = ConvertUtil.GetValueDouble(x1);
                        var d2 = ConvertUtil.GetValueDouble(y1);
                        if (double.IsNaN(d1))
                        {
                            d1 = double.MaxValue;
                        }
                        if (double.IsNaN(d2))
                        {
                            d2 = double.MaxValue;
                        }
                        ret = d1 < d2 ? -1 : (d1 > d2 ? 1 : 0);
                    }
                    else if (isNumX == false && isNumY == false)   //String Compare
                    {
                        var s1 = x1 == null ? "" : x1.ToString();
                        var s2 = y1 == null ? "" : y1.ToString();
                        ret = string.Compare(s1, s2, StringComparison.CurrentCulture);
                    }
                    else
                    {
                        ret = isNumX ? -1 : 1;
                    }
                    if (ret != 0) return ret * (descending[i] ? -1 : 1);
                }
                return 0;
            }
        }
        /// <summary>
        /// Sort the range by value of the first column, Ascending.
        /// </summary>
        public void Sort()
        {
            Sort(new int[] { 0 }, new bool[] { false });
        }
        /// <summary>
        /// Sort the range by value of the supplied column, Ascending.
        /// <param name="column">The column to sort by within the range. Zerobased</param>
        /// <param name="descending">Descending if true, otherwise Ascending. Default Ascending. Zerobased</param>
        /// </summary>
        public void Sort(int column, bool descending = false)
        {
            Sort(new int[] { column }, new bool[] { descending });
        }
        /// <summary>
        /// Sort the range by value
        /// </summary>
        /// <param name="columns">The column(s) to sort by within the range. Zerobased</param>
        /// <param name="descending">Descending if true, otherwise Ascending. Default Ascending. Zerobased</param>
        /// <param name="culture">The CultureInfo used to compare values. A null value means CurrentCulture</param>
        /// <param name="compareOptions">String compare option</param>
        public void Sort(int[] columns, bool[] descending = null, CultureInfo culture = null, CompareOptions compareOptions = CompareOptions.None)
        {
            if (columns == null)
            {
                columns = new int[] { 0 };
            }
            var cols = _toCol - _fromCol + 1;
            foreach (var c in columns)
            {
                if (c > cols - 1 || c < 0)
                {
                    throw (new ArgumentException("Can not reference columns outside the boundries of the range. Note that column reference is zero-based within the range"));
                }
            }
            var e = new CellsStoreEnumerator<ExcelCoreValue>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
            var l = new List<SortItem<ExcelCoreValue>>();
            SortItem<ExcelCoreValue> item = new SortItem<ExcelCoreValue>();

            while (e.Next())
            {
                if (l.Count == 0 || l[l.Count - 1].Row != e.Row)
                {
                    item = new SortItem<ExcelCoreValue>() { Row = e.Row, Items = new ExcelCoreValue[cols] };
                    l.Add(item);
                }
                item.Items[e.Column - _fromCol] = e.Value;
            }

            if (descending == null)
            {
                descending = new bool[columns.Length];
                for (int i = 0; i < columns.Length; i++)
                {
                    descending[i] = false;
                }
            }

            var comp = new Comp();
            comp.columns = columns;
            comp.descending = descending;
            comp.cultureInfo = culture ?? CultureInfo.CurrentCulture;
            comp.compareOptions = compareOptions;
            l.Sort(comp);

            var flags = GetItems(_worksheet._flags, _fromRow, _fromCol, _toRow, _toCol);
            var formulas = GetItems(_worksheet._formulas, _fromRow, _fromCol, _toRow, _toCol);
            var hyperLinks = GetItems(_worksheet._hyperLinks, _fromRow, _fromCol, _toRow, _toCol);
            var comments = GetItems(_worksheet._commentsStore, _fromRow, _fromCol, _toRow, _toCol);
            //Sort the values and styles.
            _worksheet._values.Clear(_fromRow, _fromCol, _toRow - _fromRow + 1, cols);
            for (var r = 0; r < l.Count; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var row = _fromRow + r;
                    var col = _fromCol + c;
                    _worksheet._values.SetValueSpecial(row, col, SortSetValue, l[r].Items[c]);
                    var addr = GetAddress(l[r].Row, _fromCol + c);
                    //Move flags
                    if (flags.ContainsKey(addr))
                    {
                        _worksheet._flags.SetValue(row, col, flags[addr]);
                    }
                    //Move formulas
                    if (formulas.ContainsKey(addr))
                    {
                        _worksheet._formulas.SetValue(row, col, formulas[addr]);
                        if (formulas[addr] is int)
                        {
                            int sfIx = (int)formulas[addr];
                            var startAddr = new ExcelAddress(Worksheet._sharedFormulas[sfIx].Address);
                            var f = Worksheet._sharedFormulas[sfIx];
                            if (startAddr._fromRow > row)
                            {
                                f.Formula = ExcelCellBase.TranslateFromR1C1(ExcelCellBase.TranslateToR1C1(f.Formula, f.StartRow, f.StartCol), row, f.StartCol);
                                f.StartRow = row;
                                f.Address = ExcelCellBase.GetAddress(row, startAddr._fromCol, startAddr._toRow, startAddr._toCol);
                            }
                            else if (startAddr._toRow < row)
                            {
                                f.Address = ExcelCellBase.GetAddress(startAddr._fromRow, startAddr._fromCol, row, startAddr._toCol);
                            }
                        }
                    }

                    //Move hyperlinks
                    if (hyperLinks.ContainsKey(addr))
                    {
                        _worksheet._hyperLinks.SetValue(row, col, hyperLinks[addr]);
                    }

                    //Move comments
                    if (comments.ContainsKey(addr))
                    {
                        var i = comments[addr];
                        _worksheet._commentsStore.SetValue(row, col, i);
                        var comment = _worksheet._comments[i];
                        comment.Reference = GetAddress(row, col);
                    }
                }
            }
        }

        private static Dictionary<string, T> GetItems<T>(CellStore<T> store, int fromRow, int fromCol, int toRow, int toCol)
        {
            var e = new CellsStoreEnumerator<T>(store, fromRow, fromCol, toRow, toCol);
            var l = new Dictionary<string, T>();
            while (e.Next())
            {
                l.Add(e.CellAddress, e.Value);
            }
            return l;
        }

        private static void SortSetValue(List<ExcelCoreValue> list, int index, object value)
        {
            var v = (ExcelCoreValue)value;
            list[index] = new ExcelCoreValue { _value = v._value, _styleId = v._styleId };
        }
    }
}
