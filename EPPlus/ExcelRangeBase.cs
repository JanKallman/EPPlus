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
		private delegate void _changeProp(_setValue method, object value);
		private delegate void _setValue(object value, int row, int col);
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
        //private class CopiedFlag
        //{
        //    internal int Row { get; set; }
        //    internal int Column { get; set; }
        //    internal Byte Flag { get; set; }
        //}
        #region Constructors
		internal ExcelRangeBase(ExcelWorksheet xlWorksheet)
		{
			_worksheet = xlWorksheet;
			_ws = _worksheet.Name;
            _workbook = _worksheet.Workbook;
            this.AddressChange += new EventHandler(ExcelRangeBase_AddressChange);
			SetDelegate();
		}

        void ExcelRangeBase_AddressChange(object sender, EventArgs e)
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
            this.AddressChange += new EventHandler(ExcelRangeBase_AddressChange);
            SetDelegate();
		}
		internal ExcelRangeBase(ExcelWorkbook wb, ExcelWorksheet xlWorksheet, string address, bool isName) :
			base(xlWorksheet == null ? "" : xlWorksheet.Name, address, isName)
		{
            SetRCFromTable(wb._package, null);
            _worksheet = xlWorksheet;
			_workbook = wb;
			if (string.IsNullOrEmpty(_ws)) _ws = (xlWorksheet == null ? null : xlWorksheet.Name);
            this.AddressChange += new EventHandler(ExcelRangeBase_AddressChange);
            SetDelegate();
		}
        ~ExcelRangeBase()
        {
            this.AddressChange -= new EventHandler(ExcelRangeBase_AddressChange);
        }
		#endregion
		#region Set Value Delegates
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
		/// <param name="valueMethod"></param>
		/// <param name="value"></param>
		private void SetUnknown(_setValue valueMethod, object value)
		{
			//Address is not set use, selected range
			if (_fromRow == -1)
			{
				SetToSelectedRange();
			}
			SetDelegate();
			_changePropMethod(valueMethod, value);
		}
		/// <summary>
		/// Set a single cell
		/// </summary>
		/// <param name="valueMethod"></param>
		/// <param name="value"></param>
		private void SetSingle(_setValue valueMethod, object value)
		{
			valueMethod(value, _fromRow, _fromCol);
		}
		/// <summary>
		/// Set a range
		/// </summary>
		/// <param name="valueMethod"></param>
		/// <param name="value"></param>
		private void SetRange(_setValue valueMethod, object value)
		{
			SetValueAddress(this, valueMethod, value);
		}
		/// <summary>
		/// Set a multirange (A1:A2,C1:C2)
		/// </summary>
		/// <param name="valueMethod"></param>
		/// <param name="value"></param>
		private void SetMultiRange(_setValue valueMethod, object value)
		{
			SetValueAddress(this, valueMethod, value);
			foreach (var address in Addresses)
			{
				SetValueAddress(address, valueMethod, value);
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
				for (int col = address.Start.Column; col <= address.End.Column; col++)
				{
					for (int row = address.Start.Row; row <= address.End.Row; row++)
					{
						valueMethod(value, row, col);
					}
				}
			}
		}
		#endregion
		#region Set property methods
		private void Set_StyleID(object value, int row, int col)
		{
            _worksheet._styles.SetValue(row, col, (int)value);
		}
		private void Set_StyleName(object value, int row, int col)
		{
			//_worksheet.Cell(row, col).SetNewStyleName(value.ToString(), _styleID);
            _worksheet._styles.SetValue(row, col, _styleID);
		}
		private void Set_Value(object value, int row, int col)
		{
			//ExcelCell c = _worksheet.Cell(row, col);
            var sfi = _worksheet._formulas.GetValue(row, col);
            if (sfi is int)
            {
                SplitFormulas(_worksheet.Cells[row, col]);                
            }
            if (sfi != null) _worksheet._formulas.SetValue(row, col, string.Empty);
			_worksheet._values.SetValue(row, col, value);
		}
		private void Set_Formula(object value, int row, int col)
		{
			//ExcelCell c = _worksheet.Cell(row, col);
            var f = _worksheet._formulas.GetValue(row, col);
            if (f is int && (int)f >= 0) SplitFormulas(_worksheet.Cells[row, col]);

			string formula = (value == null ? string.Empty : value.ToString());
			if (formula == string.Empty)
			{
                _worksheet._formulas.SetValue(row, col, string.Empty);
			}
			else
			{
				if (formula[0] == '=') value = formula.Substring(1, formula.Length - 1); // remove any starting equalsign.
                _worksheet._formulas.SetValue(row, col, formula);
                _worksheet._values.SetValue(row, col, null);
            }
		}
		/// <summary>
		/// Handles shared formulas
		/// </summary>
		/// <param name="value">The  formula</param>
		/// <param name="address">The address of the formula</param>
		/// <param name="IsArray">If the forumla is an array formula.</param>
		private void Set_SharedFormula(string value, ExcelAddress address, bool IsArray)
		{
			if (_fromRow == 1 && _fromCol == 1 && _toRow == ExcelPackage.MaxRows && _toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
			{
				throw (new InvalidOperationException("Can't set a formula for the entire worksheet"));
			}
			else if (address.Start.Row == address.End.Row && address.Start.Column == address.End.Column && !IsArray)             //is it really a shared formula? Arrayformulas can be one cell only
			{
				//Nope, single cell. Set the formula
				Set_Formula(value, address.Start.Row, address.Start.Column);
				return;
			}
			//RemoveFormuls(address);
			CheckAndSplitSharedFormula(address);
			ExcelWorksheet.Formulas f = new ExcelWorksheet.Formulas(SourceCodeTokenizer.Default);
			f.Formula = value;
			f.Index = _worksheet.GetMaxShareFunctionIndex(IsArray);
			f.Address = address.FirstAddress;
			f.StartCol = address.Start.Column;
			f.StartRow = address.Start.Row;
			f.IsArray = IsArray;

			_worksheet._sharedFormulas.Add(f.Index, f);
            //_worksheet.Cell(address.Start.Row, address.Start.Column).SharedFormulaID = f.Index;
            //_worksheet.Cell(address.Start.Row, address.Start.Column).Formula = value;

			for (int col = address.Start.Column; col <= address.End.Column; col++)
			{
				for (int row = address.Start.Row; row <= address.End.Row; row++)
				{
					//_worksheet.Cell(row, col).SharedFormulaID = f.Index;
                    _worksheet._formulas.SetValue(row, col, f.Index);
                    _worksheet._values.SetValue(row, col, null);
				}
			}
		}
		private void Set_HyperLink(object value, int row, int col)
		{
			//_worksheet.Cell(row, col).Hyperlink = value as Uri;
            if (value is Uri)
            {
                _worksheet._hyperLinks.SetValue(row, col, (Uri)value);

                if (value is ExcelHyperLink)
                {
                    _worksheet._values.SetValue(row, col, ((ExcelHyperLink)value).Display);
                }
                else
                {
                   _worksheet._values.SetValue(row, col, ((Uri)value).OriginalString);
                }                    
            }
            else
            {
                _worksheet._hyperLinks.SetValue(row, col, (Uri)null);
                _worksheet._values.SetValue(row, col, (Uri)null);
            }
        }
		private void Set_IsRichText(object value, int row, int col)
		{
			//_worksheet.Cell(row, col).IsRichText = (bool)value;
            _worksheet._flags.SetFlagValue(row, col, (bool)value, CellFlags.RichText);
		}
		private void Exists_Comment(object value, int row, int col)
		{
			ulong cellID = GetCellID(_worksheet.SheetID, row, col);
			if (_worksheet.Comments._comments.ContainsKey(cellID))
			{
				throw (new InvalidOperationException(string.Format("Cell {0} already contain a comment.", new ExcelCellAddress(row, col).Address)));
			}

		}
		private void Set_Comment(object value, int row, int col)
		{
			string[] v = (string[])value;
			Worksheet.Comments.Add(new ExcelRangeBase(_worksheet, GetAddress(_fromRow, _fromCol)), v[0], v[1]);
			//   _worksheet.Cell(row, col).Comment = comment;
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
		#region Public Properties
		/// <summary>
		/// The styleobject for the range.
		/// </summary>
		public ExcelStyle Style
		{
			get
			{
				IsRangeValid("styling");
                int s=0;
                if(!_worksheet._styles.Exists(_fromRow,_fromCol, ref s)) //Cell exists
                {
                    if(!_worksheet._styles.Exists(_fromRow,0, ref s)) //No, check Row style
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
				int  xfId;
                if (_fromRow == 1 && _toRow == ExcelPackage.MaxRows)
                {
                    xfId=GetColumnStyle(_fromCol);
                }
                else if (_fromCol == 1 && _toCol == ExcelPackage.MaxColumns)
                {
                    xfId = 0;
                    if (!_worksheet._styles.Exists(_fromRow, 0, ref xfId))
                    {
                        xfId = GetColumnStyle(_fromCol);
                    }
                }
                else
                {
                    xfId = 0;
                    if(!_worksheet._styles.Exists(_fromRow, _fromCol, ref xfId))
                    {
                        if (!_worksheet._styles.Exists(_fromRow, 0, ref xfId))
                        {
                            xfId = GetColumnStyle(_fromCol);
                        }
                    }
                }
                int nsID;
                if (xfId <= 0)
                {
                    nsID=Style.Styles.CellXfs[0].XfId;
                }
                else
                {
                    nsID=Style.Styles.CellXfs[xfId].XfId;
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
					//Get the startcolumn
					//ulong colID = ExcelColumn.GetColumnID(_worksheet.SheetID, column);
                    var c = _worksheet.GetValue(0, _fromCol);
                    if (c==null)
					{
                        column = _worksheet.Column(_fromCol);
                        //if (_worksheet._values.PrevCell(ref row, ref col))
                        //{
                        //    var prevCol = (ExcelColumn)_worksheet._values.GetValue(row, col);
                        //    column = prevCol.Clone(_worksheet, column);
                        //    prevCol.ColumnMax = column - 1;
                        //}
					}
					else
					{
                        column = (ExcelColumn)c;
					}

                    column.StyleName = value;
                    column.StyleID = _styleID;

                    //var index = _worksheet._columns.IndexOf(colID);
                    var cols = new CellsStoreEnumerator<object>(_worksheet._values, 0, _fromCol + 1, 0, _toCol);
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

                            if (cols.Value == null)
                            {
                                break;
                            }
                            else
                            {
                                var nextCol = (ExcelColumn)cols.Value;
                                if(column.ColumnMax < nextCol.ColumnMax-1)
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
                    //if (column.ColumnMin == column)
                    //{
                    //    column.ColumnMax = _toCol;
                    //}
                    //else if (column._columnMax < _toCol)
                    //{
                    //    var newCol = _worksheet.Column(column._columnMax + 1) as ExcelColumn;
                    //    newCol._columnMax = _toCol;

                    //    newCol._styleID = _styleID;
                    //    newCol._styleName = value;
                    //}
                    if (_fromCol == 1 && _toCol == ExcelPackage.MaxColumns) //FullRow
                    {
                        var rows = new CellsStoreEnumerator<object>(_worksheet._values, 1, 0, ExcelPackage.MaxRows, 0);
                        rows.Next();
                        while(rows.Value!=null)
                        {
                            _worksheet._styles.SetValue(rows.Row, 0, _styleID);
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
                            _worksheet._styles.SetValue(r, c, _styleID);
                        }
                    }
                }
                else //Only set name on created cells. (uncreated cells is set on full row or full column).
                {
                    var cells = new CellsStoreEnumerator<object>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
                    while (cells.Next())
                    {
                        _worksheet._styles.SetValue(cells.Row, cells.Column, _styleID);
                    }
                }
                //_changePropMethod(Set_StyleName, value);
			}
		}

        private int GetColumnStyle(int col)
        {
            object c=null;
            if (_worksheet._values.Exists(0, col, ref c))
            {
                return (c as ExcelColumn).StyleID;
            }
            else
            {
                int row = 0;
                if (_worksheet._values.PrevCell(ref row, ref col))
                {
                    var column=_worksheet._values.GetValue(row,col) as ExcelColumn;
                    if(column.ColumnMax>=col)
                    {
                        return _worksheet._styles.GetValue(row, col);
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
                int s=0;
                if(!_worksheet._styles.Exists(_fromRow, _fromCol, ref s))
                {
                    if (!_worksheet._styles.Exists(_fromRow, 0, ref s))
                    {
                        s = _worksheet._styles.GetValue(0, _fromCol);
                    }
                }
                return s;
			}
			set
			{
				_changePropMethod(Set_StyleID, value);
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
					_changePropMethod(Set_Value, value);
				}
			}
		}

        private bool IsInfinityValue(object value)
        {
            double? valueAsDouble = value as double?;

            if(valueAsDouble.HasValue && 
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
					if (_worksheet._values.Exists(row,col))
					{
                        if (_worksheet._flags.GetFlagValue(row, col, CellFlags.RichText))
						{
							v[row - addr._fromRow, col - addr._fromCol] = GetRichText(row, col).Text;
						}
						else
						{
							v[row - addr._fromRow, col - addr._fromCol] = _worksheet._values.GetValue(row, col);
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

			if (addr._fromCol == fromRow && addr._fromCol == addr._fromCol && addr._toRow == toRow && addr._toCol == _toCol)
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
				return _worksheet._values.GetValue(_fromRow, _fromCol);
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
		/// Note: Cells containing formulas are ignored since EPPlus don't have a calculation engine.
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
				afAddr.Add(new ExcelAddressBase(    _worksheet.AutoFilterAddress._fromRow,
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
            
			using (var b = new Bitmap(1, 1))
			{
				using (var g = Graphics.FromImage(b))
				{
					var normalSize = (float)Math.Truncate(g.MeasureString("00", nfont).Width - g.MeasureString("0", nfont).Width);
					g.PageUnit = GraphicsUnit.Pixel;
					foreach (var cell in this)
					{
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

						//Truncate(({pixels}-5)/{Maximum Digit Width} * 100+0.5)/100

                        var size = g.MeasureString(cell.TextForWidth, f);
                        double width;
                        double r = styles.CellXfs[cell.StyleID].TextRotation;
                        if (r <= 0 )
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
				}
			}
			_worksheet.Drawings.AdjustWidth(drawWidths);
			_worksheet._package.DoAdjustDrawings = doAdjust;
		}

        private void SetMinWidth(double minimumWidth, int fromCol, int toCol)
        {
            var iterator = new CellsStoreEnumerator<object>(_worksheet._values, 0, fromCol, 0, toCol);
            var prevCol = fromCol;
            foreach (ExcelColumn col in iterator)
            {
                col.Width = minimumWidth;
                if (_worksheet.DefaultColWidth > minimumWidth && col.ColumnMin > prevCol)
                {
                    var newCol = _worksheet.Column(prevCol);
                    newCol.ColumnMax = col.ColumnMin - 1;
                    newCol.Width = minimumWidth;
                }
                prevCol = col.ColumnMax + 1;
            }
            if (_worksheet.DefaultColWidth > minimumWidth && prevCol<toCol)
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
			if (v is decimal || v.GetType().IsPrimitive)
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
					return date.ToString(format, nf.Culture);
				}
			}
			else if (v is DateTime)
			{
				if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
				{
					return ((DateTime)v).ToString(format, nf.Culture);
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
					return new DateTime(((TimeSpan)v).Ticks).ToString(format, nf.Culture);
				}
				else
				{
					double d = (new DateTime(((TimeSpan)v).Ticks)).ToOADate();
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
					if(value==null || value.Trim()=="")
                    {
                        //Set the cells to null
                        Value = null;
                    }                    
                    else if (_fromRow == _toRow && _fromCol == _toCol)
					{
						Set_Formula(value, _fromRow, _fromCol);
					}
					else
					{
						Set_SharedFormula(value, this, false);
						if (Addresses != null)
						{
							foreach (var address in Addresses)
							{
								Set_SharedFormula(value, address, false);
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
					Set_SharedFormula(ExcelCellBase.TranslateFromR1C1(value, _fromRow, _fromCol), this, false);
				}
				else
				{
					Set_SharedFormula(ExcelCellBase.TranslateFromR1C1(value, _fromRow, _fromCol), new ExcelAddress(WorkSheet, FirstAddress), false);
					foreach (var address in Addresses)
					{
						Set_SharedFormula(ExcelCellBase.TranslateFromR1C1(value, address.Start.Row, address.Start.Column), address, false);
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
				_changePropMethod(Set_HyperLink, value);
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
                        if(_worksheet.MergedCells[row, col]==null)
                        {
                            return false;
                        }
                        //if (!_worksheet._flags.GetFlagValue(row, col, CellFlags.Merged))
                        //{
                        //    return false;
                        //}
					}
				}
				return true;
			}
			set
			{
				IsRangeValid("merging");
				//SetMerge(value, FirstAddress);
			    if (value)
			    {
			        _worksheet.MergedCells.Add(new ExcelAddressBase(FirstAddress), true);
			        if (Addresses != null)
			        {
			            foreach (var address in Addresses)
			            {
			                _worksheet.MergedCells.Add(address, true);
			                //SetMerge(value, address._address);
			            }
			        }
			    }
			    else
			    {
			        _worksheet.MergedCells.Clear(this);
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

        //private void SetMerge(bool value, string address)
        //{
        //    if (!value)
        //    {
        //        if (_worksheet.MergedCells.List.Contains(address))
        //        {
        //            SetCellMerge(false, address);
        //            _worksheet.MergedCells.List.Remove(address);
        //        }
        //        else if (!CheckMergeDiff(false, address))
        //        {
        //            throw (new Exception("Range is not fully merged.Specify the exact range"));
        //        }
        //    }
        //    else
        //    {
        //        if (CheckMergeDiff(false, address))
        //        {
        //            SetCellMerge(true, address);
        //            _worksheet.MergedCells.List.Add(address);
        //        }
        //        else
        //        {
        //            if (!_worksheet.MergedCells.List.Contains(address))
        //            {
        //                throw (new Exception("Cells are already merged"));
        //            }
        //        }
        //    }
        //}
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
				_worksheet.AutoFilterAddress = this;
				if (_worksheet.Names.ContainsKey("_xlnm._FilterDatabase"))
				{
					_worksheet.Names.Remove("_xlnm._FilterDatabase");
				}
				var result = _worksheet.Names.Add("_xlnm._FilterDatabase", this);
				result.IsNameHidden = true;
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
				return _worksheet._flags.GetFlagValue(_fromRow, _fromCol,CellFlags.RichText);
			}
			set
			{
				_changePropMethod(Set_IsRichText, value);
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
		ExcelRichTextCollection _rtc = null;
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
            var v = _worksheet._values.GetValue(row, col);
            var isRt = _worksheet._flags.GetFlagValue(row, col, CellFlags.RichText);
            if (v != null)
			{
				if (isRt)
				{
                    XmlHelper.LoadXmlSafe(xml, "<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" >" + v.ToString() + "</d:si>", Encoding.UTF8);
				}
				else
				{
					xml.LoadXml("<d:si xmlns:d=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ><d:r><d:t>" + SecurityElement.Escape(v.ToString()) + "</d:t></d:r></d:si>");
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
				ulong cellID = GetCellID(_worksheet.SheetID, _fromRow, _fromCol);
				if (_worksheet.Comments._comments.ContainsKey(cellID))
				{
					return _worksheet._comments._comments[cellID] as ExcelComment;
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
		public string FullAddress
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
				string fullAddress = GetFullAddress(wbwsRef, GetAddress(_fromRow, _fromCol, _toRow, _toCol, true));
				if (Addresses != null)
				{
					foreach (var a in Addresses)
					{
						fullAddress += "," + GetFullAddress(wbwsRef, GetAddress(a.Start.Row, a.Start.Column, a.End.Row, a.End.Column, true)); ;
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
                string fullAddress = GetFullAddress(wbwsRef, GetAddress(_fromRow, _fromCol, _toRow, _toCol, true), false);
                if (Addresses != null)
                {
                    foreach (var a in Addresses)
                    {
                        fullAddress += "," + GetFullAddress(wbwsRef, GetAddress(a.Start.Row, a.Start.Column, a.End.Row, a.End.Column, true),false); ;
                    }
                }
                return fullAddress;
            }
        }
		#endregion
		#region Private Methods
        ///// <summary>
        ///// Check if the range is partly merged
        ///// </summary>
        ///// <param name="startValue">the starting value</param>
        ///// <param name="address">the address</param>
        ///// <returns></returns>
        //private bool CheckMergeDiff(bool startValue, string address)
        //{
        //    ExcelAddress a = new ExcelAddress(address);
        //    for (int col = a.column; col <= a._toCol; col++)
        //    {
        //        for (int row = a._fromRow; row <= a._toRow; row++)
        //        {
        //            if (_worksheet._flags.GetFlagValue(row, col, CellFlags.Merged) != startValue)
        //            {
        //                return false;
        //            }
        //        }
        //    }
        //    return true;
        //}
        ///// <summary>
        ///// Set the merge flag for the range
        ///// </summary>
        ///// <param name="value"></param>
        ///// <param name="address"></param>
        //internal void SetCellMerge(bool value, string address)
        //{
        //    ExcelAddress a = new ExcelAddress(address);
        //    for (int col = a.column; col <= a._toCol; col++)
        //    {
        //        for (int row = a._fromRow; row <= a._toRow; row++)
        //        {
        //            _worksheet._flags.SetFlagValue(row, col,value,CellFlags.Merged);
        //        }
        //    }
        //}
		/// <summary>
		/// Set the value without altering the richtext property
		/// </summary>
		/// <param name="value">the value</param>
		internal void SetValueRichText(object value)
		{
			if (_fromRow == 1 && _fromCol == 1 && _toRow == ExcelPackage.MaxRows && _toCol == ExcelPackage.MaxColumns)  //Full sheet (ex ws.Cells.Value=0). Set value for A1 only to avoid hanging 
			{
				//_worksheet.Cell(1, 1).SetValueRichText(value);
                SetValue(value, 1, 1);
			}
			else
			{
                //for (int col = _fromCol; col <= _toCol; col++)
                //{
                //    for (int row = _fromRow; row <= _toRow; row++)
                //    {
						//_worksheet.Cell(row, col).SetValueRichText(value);
                        SetValue(value, _fromRow,_fromCol);
                    //}
				//}
			}
		}

        private void SetValue(object value, int row, int col)
        {
            _worksheet.SetValue(row, col, value);
           // if (value is string) _worksheet._types.SetValue(row, col, "S"); else _worksheet._types.SetValue(row, col, "");
            _worksheet._formulas.SetValue(row, col, "");
        }
		/// <summary>
		/// Removes a shared formula
		/// </summary>
        //private void RemoveFormuls(ExcelAddress address)
        //{
        //    List<int> removed = new List<int>();
        //    int fFromRow, fFromCol, fToRow, fToCol;
        //    foreach (int index in _worksheet._sharedFormulas.Keys)
        //    {
        //        ExcelWorksheet.Formulas f = _worksheet._sharedFormulas[index];
        //        ExcelCellBase.GetRowColFromAddress(f.Address, out fFromRow, out fFromCol, out fToRow, out fToCol);
        //        if (((fFromCol >= address.Start.Column && fFromCol <= address.End.Column) ||
        //             (fToCol >= address.Start.Column && fToCol <= address.End.Column)) &&
        //             ((fFromRow >= address.Start.Row && fFromRow <= address.End.Row) ||
        //             (fToRow >= address.Start.Row && fToRow <= address.End.Row)))
        //        {
        //            for (int col = fFromCol; col <= fToCol; col++)
        //            {
        //                for (int row = fFromRow; row <= fToRow; row++)
        //                {
        //                    _worksheet._formulas.SetValue(row, col, int.MinValue);
        //                }
        //            }
        //            removed.Add(index);
        //        }
        //    }
        //    foreach (int index in removed)
        //    {
        //        _worksheet._sharedFormulas.Remove(index);
        //    }
        //}
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
			if (isText && (Format.DataTypes == null || Format.DataTypes.Length < col)) return v;

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
					return v;
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

					default:
						return v;

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
	                _worksheet._values.SetValue(row, col++, Reader.GetName(i));
	            }
	            row++;
	            col = _fromCol;
	        }
	        while(Reader.Read())
	        {
	            for (int i = 0; i < fieldCount; i++)
	            {
	                _worksheet._values.SetValue(row, col++, Reader.GetValue(i));
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
            if (rows >= 0 && Table.Columns.Count>0)
			{
                var tbl = _worksheet.Tables.Add(new ExcelAddressBase(_fromRow, _fromCol, _fromRow + rows - 1, _fromCol + Table.Columns.Count-1), Table.TableName);
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

			int col = _fromCol, row = _fromRow;
			if (PrintHeaders)
			{
				foreach (DataColumn dc in Table.Columns)
				{
                    // If no caption is set, the ColumnName property is called implicitly.
					_worksheet._values.SetValue(row, col++, dc.Caption);
				}
				row++;
				col = _fromCol;
			}
			foreach (DataRow dr in Table.Rows)
			{
				foreach (object value in dr.ItemArray)
				{
                    if (value != null && value != DBNull.Value && !string.IsNullOrEmpty(value.ToString()))
                    {
                        _worksheet._values.SetValue(row, col++, value);
                    }
                    else
                    {
                        col++;
                    }
                }
				row++;
				col = _fromCol;
			}
            return _worksheet.Cells[_fromRow, _fromCol, (row == _fromRow ? _fromRow : row - 1), _fromCol + Table.Columns.Count - 1];
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

			int column = _fromCol, row = _fromRow;

			foreach (var rowData in Data)
			{
				column = _fromCol;
				foreach (var cellData in rowData)
				{
					_worksheet._values.SetValue(row, column, cellData);
					column += 1;
				}
				row += 1;
			}
			return _worksheet.Cells[_fromRow, _fromCol, row - 1, column - 1];
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
			if (Members == null)
			{
				Members = type.GetProperties(memberFlags);
			}
			else
			{
				foreach (var t in Members)
				{
                    if (t.DeclaringType!=null && t.DeclaringType != type && !t.DeclaringType.IsSubclassOf(type))
					{
						throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T");
					}
				}
			}

			int col = _fromCol, row = _fromRow;
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
                            t.GetCustomAttributes(typeof (DisplayNameAttribute), false).FirstOrDefault() as
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
                    _worksheet._values.SetValue(row, col++, header);
				}
				row++;
			}

			if (Members.Length == 0)
			{
				foreach (var item in Collection)
				{
					_worksheet.Cells[row++, col].Value = item;
				}
			}
			else
			{
				foreach (var item in Collection)
				{
					col = _fromCol;
                    if (item is string || item is decimal || item is DateTime || item.GetType().IsPrimitive)
                    {
                        _worksheet.Cells[row, col++].Value = item;
                    }
                    else
                    {
                        foreach (var t in Members)
                        {
                            if (t is PropertyInfo)
                            {
                                _worksheet.Cells[row, col++].Value = ((PropertyInfo)t).GetValue(item, null);
                            }
                            else if (t is FieldInfo)
                            {
                                _worksheet.Cells[row, col++].Value = ((FieldInfo)t).GetValue(item);
                            }
                            else if (t is MethodInfo)
                            {
                                _worksheet.Cells[row, col++].Value = ((MethodInfo)t).Invoke(item, null);
                            }
                        }
                    }
					row++;
				}
			}

            if (_fromRow == row-1 && PrintHeaders)
            {
                row++;
            }

            var r = _worksheet.Cells[_fromRow, _fromCol, row - 1, Members.Length==0 ? col : col - 1];

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

            string splitRegex = String.Format("{0}(?=(?:[^{1}]*{1}[^{1}]*{1})*[^{1}]*$)", Format.EOL, Format.TextQualifier);
            string[] lines = Regex.Split(Text, splitRegex);
            int row = _fromRow;
            int col = _fromCol;
            int maxCol = col;
            int lineNo = 1;
            foreach (string line in lines)
            {
                if (lineNo > Format.SkipLinesBeginning && lineNo <= lines.Length - Format.SkipLinesEnd)
                {
                    col = _fromCol;
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
                                    _worksheet.SetValue(row, col, ConvertData(Format, v, col - _fromCol, isText));
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
                    if (QCount > 1)
                    {
                        v += new string(Format.TextQualifier, QCount / 2);
                    }
                    if (lineQCount % 2 == 1)
                        throw (new Exception(string.Format("Text delimiter is not closed in line : {0}", line)));

                    _worksheet._values.SetValue(row, col, ConvertData(Format, v, col - _fromCol, isText));
                    if (col > maxCol) maxCol = col;
                    row++;
                }
                lineNo++;
            }
            return _worksheet.Cells[_fromRow, _fromCol, row - 1, maxCol];
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
		/// Get the strongly typed value of the cell.
		/// </summary>
		/// <typeparam name="T">The type</typeparam>
		/// <returns>The value. If the value can't be converted to the specified type, the default value will be returned</returns>
		public T GetValue<T>()
		{
			return _worksheet.GetTypedValue<T>(Value);
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
			//Check if any comments exists in the range and throw an exception
			_changePropMethod(Exists_Comment, null);
			//Create the comments
			_changePropMethod(Set_Comment, new string[] { Text, Author });

			return _worksheet.Comments[new ExcelCellAddress(_fromRow, _fromCol)];
		}

        ///// <summary>
        ///// Copies the range of cells to an other range
        ///// </summary>
        ///// <param name="Destination">The start cell where the range will be copied.</param>
        public void Copy(ExcelRangeBase Destination)
        {
            bool sameWorkbook = Destination._worksheet.Workbook == _worksheet.Workbook;
            ExcelStyles sourceStyles = _worksheet.Workbook.Styles,
                        styles = Destination._worksheet.Workbook.Styles;
            Dictionary<int, int> styleCashe = new Dictionary<int, int>();

            //Clear all existing cells; 
            int toRow = _toRow - _fromRow + 1,
                toCol = _toCol - _fromCol + 1;

            string s = "";
            int i=0;
            object o = null;
            byte flag=0;
            Uri hl = null;
            ExcelComment comment=null;

            var cse = new CellsStoreEnumerator<object>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
            var copiedValue = new List<CopiedCell>();
            while (cse.Next())
            {
                var row=cse.Row;
                var col = cse.Column;       //Issue 15070
                var cell = new CopiedCell
                {
                    Row = Destination._fromRow + (row - _fromRow),
                    Column = Destination._fromCol + (col - _fromCol),
                    Value=cse.Value
                };

                //Destination._worksheet._values.SetValue(row, col, cse.Value);

                if (_worksheet._types.Exists(row, col, ref s))
                {
                    //Destination._worksheet._types.SetValue(row, col,s);
                    cell.Type=s;
                }

                if (_worksheet._formulas.Exists(row, col, ref o))
                {
                    if (o is int)
                    {
                       // Destination._worksheet._formulas.SetValue(row, col, _worksheet.GetFormula(cse.Row, cse.Column));    //Shared formulas, set the formula per cell to simplify
                        cell.Formula=_worksheet.GetFormula(cse.Row, cse.Column);
                    }
                    else
                    {
                        //Destination._worksheet._formulas.SetValue(row, col, o);
                        cell.Formula=o;
                    }
                }
                if(_worksheet._styles.Exists(row, col, ref i))
                {
                    if (sameWorkbook)
                    {
                        //Destination._worksheet._styles.SetValue(row, col, i);
                        cell.StyleID=i;
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
                        //Destination._worksheet._styles.SetValue(row, col, i);
                        cell.StyleID=i;
                    }
                }
                
                if (_worksheet._hyperLinks.Exists(row, col, ref hl))
                {
                    //Destination._worksheet._hyperLinks.SetValue(row, col, hl);
                    cell.HyperLink=hl;
                }

                if(_worksheet._commentsStore.Exists(row, col, ref comment))
                {
                    cell.Comment=comment;
                }

                if (_worksheet._flags.Exists(row, col, ref flag))
                {
                    cell.Flag = flag;
                }
                copiedValue.Add(cell);
            }

            //Copy styles with no cell value
            var cses = new CellsStoreEnumerator<int>(_worksheet._styles, _fromRow, _fromCol, _toRow, _toCol);
            while (cses.Next())
            {
                if (!_worksheet._values.Exists(cses.Row, cses.Column))
                {
                    var row = Destination._fromRow + (cses.Row - _fromRow);
                    var col = Destination._fromCol + (cses.Column - _fromCol);
                    var cell = new CopiedCell
                    {
                        Row = row,
                        Column = col,
                        Value = null
                    };

                    i = cses.Value;
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
                        //Destination._worksheet._styles.SetValue(row, col, i);
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
                if(!copiedMergedCells.ContainsKey(csem.Value))
                {
                    var adr = new ExcelAddress(_worksheet.Name, _worksheet.MergedCells.List[csem.Value]);
                    if(this.Collide(adr)==eAddressCollition.Inside)
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

            Destination._worksheet.MergedCells.Clear(new ExcelAddressBase(Destination._fromRow, Destination._fromCol, Destination._fromRow+toRow-1, Destination._fromCol+toCol-1));

            Destination._worksheet._values.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._formulas.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._styles.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._types.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._hyperLinks.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._flags.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
            Destination._worksheet._commentsStore.Clear(Destination._fromRow, Destination._fromCol, toRow, toCol);
           
            foreach(var cell in copiedValue)
            {
                Destination._worksheet._values.SetValue(cell.Row, cell.Column, cell.Value);

                if(cell.Type!=null)
                {
                    Destination._worksheet._types.SetValue(cell.Row, cell.Column, cell.Type);
                }

                if(cell.StyleID!=null)
                {
                    Destination._worksheet._styles.SetValue(cell.Row, cell.Column, cell.StyleID.Value);
                }

                if(cell.Formula!=null)
                {
                    cell.Formula = UpdateFormulaReferences(cell.Formula.ToString(), Destination._fromRow - _fromRow, Destination._fromCol - _fromCol, 0, 0, true);
                    Destination._worksheet._formulas.SetValue(cell.Row, cell.Column, cell.Formula);
                }
                if(cell.HyperLink!=null)
                {
                    Destination._worksheet._hyperLinks.SetValue(cell.Row, cell.Column, cell.HyperLink);
                }

                if (cell.Comment != null)
                {
                    //Destination._worksheet._commentsStore.SetValue(cell.Row, cell.Column, cell.Comment);
                }
                if (cell.Flag != 0)
                {
                    Destination._worksheet._flags.SetValue(cell.Row, cell.Column, cell.Flag);
                }
            }

            //Add merged cells
            foreach(var m in copiedMergedCells.Values)
            {
                if(m!=null)
                {
                    Destination._worksheet.MergedCells.Add(m, true);
                }
            }


            //Clone the cell
                //var copiedCell = (_worksheet._cells[GetCellID(_worksheet.SheetID, cell._fromRow, cell.column)] as ExcelCell);

                //var newCell = copiedCell.Clone(Destination._worksheet,
                //        Destination._fromRow + (copiedCell.Row - _fromRow),
                //        Destination.column + (copiedCell.Column - column));

        //        newCell.MergeId = _worksheet.GetMergeCellId(copiedCell.Row, copiedCell.Column);


        //        if (!string.IsNullOrEmpty(newCell.Formula))
        //        {
        //            newCell.Formula = ExcelCell.UpdateFormulaReferences(newCell.Formula, newCell.Row - copiedCell.Row, (newCell.Column - copiedCell.Column), 1, 1);
        //        }

        //        //If its not the same workbook we must copy the styles to the new workbook.
        //        if (!sameWorkbook)
        //        {
        //            if (styleCashe.ContainsKey(cell.StyleID))
        //            {
        //                newCell.StyleID = styleCashe[cell.StyleID];
        //            }
        //            else
        //            {
        //                newCell.StyleID = styles.CloneStyle(sourceStyles, cell.StyleID);
        //                styleCashe.Add(cell.StyleID, newCell.StyleID);
        //            }
        //        }
        //        newCells.Add(newCell);
        //        if (newCell.Merge) mergedCells.Add(newCell.CellID, newCell);
        //    }

        //    //Now clear the destination.
        //    Destination.Offset(0, 0, (_toRow - _fromRow) + 1, (_toCol - column) + 1).Clear();

        //    //And last add the new cells to the worksheet
        //    foreach (var cell in newCells)
        //    {
        //        Destination.Worksheet._cells.Add(cell);
        //    }
        //    //Add merged cells
        //    if (mergedCells.Count > 0)
        //    {
        //        List<ExcelAddressBase> mergedAddresses = new List<ExcelAddressBase>();
        //        foreach (var cell in mergedCells.Values)
        //        {
        //            if (!IsAdded(cell, mergedAddresses))
        //            {
        //                int startRow = cell.Row, startCol = cell.Column, endRow = cell.Row, endCol = cell.Column + 1;
        //                while (mergedCells.ContainsKey(ExcelCell.GetCellID(Destination.Worksheet.SheetID, endRow, endCol)))
        //                {
        //                    ExcelCell next = mergedCells[ExcelCell.GetCellID(Destination.Worksheet.SheetID, endRow, endCol)];
        //                    if (cell.MergeId != next.MergeId)
        //                    {
        //                        break;
        //                    }
        //                    endCol++;
        //                }

        //                while (IsMerged(mergedCells, Destination.Worksheet, endRow, startCol, endCol - 1, cell))
        //                {
        //                    endRow++;
        //                }

        //                mergedAddresses.Add(new ExcelAddressBase(startRow, startCol, endRow - 1, endCol - 1));
        //            }
        //        }
        //        Destination.Worksheet.MergedCells.List.AddRange((from r in mergedAddresses select r.Address));
        //    }
        //}

        //private bool IsAdded(ExcelCell cell, List<ExcelAddressBase> mergedAddresses)
        //{
        //    foreach (var address in mergedAddresses)
        //    {
        //        if (address.Collide(new ExcelAddressBase(cell.CellAddress)) == eAddressCollition.Inside)
        //        {
        //            return true;
        //        }
        //    }
        //    return false;
        //}

        //private bool IsMerged(Dictionary<ulong, ExcelCell> mergedCells, ExcelWorksheet worksheet, int row, int startCol, int endCol, ExcelCell cell)
        //{
        //    for (int col = startCol; col <= endCol; col++)
        //    {
        //        if (!mergedCells.ContainsKey(ExcelCell.GetCellID(worksheet.SheetID, row, col)))
        //        {
        //            return false;
        //        }
        //        else
        //        {
        //            ExcelCell next = mergedCells[ExcelCell.GetCellID(worksheet.SheetID, row, col)];
        //            if (cell.MergeId != next.MergeId)
        //            {
        //                return false;
        //            }
        //        }
        //    }
        //    return true;
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
			Set_SharedFormula(ArrayFormula, this, true);
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
                fromCol = Range._fromRow;
            }

            var rows = Range._toRow - fromRow + 1;
            var cols = Range._toCol - fromCol + 1;
            
            _worksheet._values.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._types.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._styles.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._formulas.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._hyperLinks.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._flags.Delete(fromRow, fromCol, rows, cols, shift);
            _worksheet._commentsStore.Delete(fromRow, fromCol, rows, cols, shift);

            //if(shift)
            //{
            //    _worksheet.AdjustFormulasRow(fromRow, rows);
            //}

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
        //int _index;
        //ulong _toCellId;
        //int _enumAddressIx;
        CellsStoreEnumerator<object> cellEnum;
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

		int _enumAddressIx = -1;
        public bool MoveNext()
		{
            if (cellEnum.Next())
            {
                return true;
            }
            else if (_addresses!=null)
            {
                _enumAddressIx++;
                if (_enumAddressIx < _addresses.Count)
                {
                    cellEnum = new CellsStoreEnumerator<object>(_worksheet._values, 
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
            cellEnum = new CellsStoreEnumerator<object>(_worksheet._values, _fromRow, _fromCol, _toRow, _toCol);
        }

        //private void GetNextIndexEnum(int fromRow, int fromCol, int toRow, int toCol)
        //{
        //    if (_index >= _worksheet._cells.Count) return;
        //    ExcelCell cell = _worksheet._cells[_index] as ExcelCell;
        //    while (cell.Column > toCol || cell.Column < fromCol)
        //    {
        //        if (cell.Column < fromCol)
        //        {
        //            _index = _worksheet._cells.IndexOf(ExcelAddress.GetCellID(_worksheet.SheetID, cell.Row, fromCol));
        //        }
        //        else
        //        {
        //            _index = _worksheet._cells.IndexOf(ExcelAddress.GetCellID(_worksheet.SheetID, cell.Row + 1, fromCol));
        //        }

        //        if (_index < 0)
        //        {
        //            _index = ~_index;
        //        }
        //        if (_index >= _worksheet._cells.Count || _worksheet._cells[_index].RangeID > _toCellId)
        //        {
        //            break;
        //        }
        //        cell = _worksheet._cells[_index] as ExcelCell;
        //    }
        //}

        //private void GetStartIndexEnum(int fromRow, int fromCol, int toRow, int toCol)
        //{
        //    _index = _worksheet._cells.IndexOf(ExcelCellBase.GetCellID(_worksheet.SheetID, fromRow, fromCol));
        //    _toCellId = ExcelCellBase.GetCellID(_worksheet.SheetID, toRow, toCol);
        //    if (_index < 0)
        //    {
        //        _index = ~_index;
        //    }
        //    _index--;
        //}
    #endregion
    }
}
