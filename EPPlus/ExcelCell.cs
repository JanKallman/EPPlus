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
 * Jan Källman		Added		10-SEP-2009
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;
namespace OfficeOpenXml
{
    internal class ExcelCell : ExcelCellBase, IExcelCell, IRangeID
    {
        [Flags]
        private enum flags
        {
            isMerged=1,
            isRichText=2,

        }
		#region Cell Private Properties
		private ExcelWorksheet _worksheet;
        private int _row;
        private int _col;
		internal string _formula="";
        internal string _formulaR1C1 = "";
        private Uri _hyperlink = null;
        string _dataType = "";
        #endregion
		#region ExcelCell Constructor
		/// <summary>
		/// A cell in the worksheet. 
		/// </summary>
		/// <param name="worksheet">A reference to the worksheet</param>
		/// <param name="row">The row number</param>
		/// <param name="col">The column number</param>
		internal ExcelCell(ExcelWorksheet worksheet, int row, int col)
		{
			if (row < 1 || col < 1)
                throw new ArgumentException("Negative row and column numbers are not allowed");
            if (row > ExcelPackage.MaxRows || col > ExcelPackage.MaxColumns)
                throw new ArgumentException("Row or column numbers are out of range");
            if (worksheet == null)
				throw new ArgumentException("Worksheet must be set to a valid reference");

			_row = row;
			_col = col;
            _worksheet = worksheet;
            if (col < worksheet._minCol) worksheet._minCol = col;
            if (col > worksheet._maxCol) worksheet._maxCol = col;
            _sharedFormulaID = int.MinValue;
            IsRichText = false;
		}
        internal ExcelCell(ExcelWorksheet worksheet, string cellAddress)
        {
            _worksheet = worksheet;
            GetRowColFromAddress(cellAddress, out _row, out _col);
            if (_col < worksheet._minCol) worksheet._minCol = _col;
            if (_col > worksheet._maxCol) worksheet._maxCol = _col;
            _sharedFormulaID = int.MinValue;
            IsRichText = false;
        }
		#endregion 
        internal ulong CellID
        {
            get
            {
                return GetCellID(_worksheet.SheetID, Row, Column);
            }
        }
        #region ExcelCell Public Properties

		/// <summary>
		/// Row number
		/// </summary>
        internal int Row { get { return _row; } set { _row = value; } }
		/// <summary>
		/// Column number
		/// </summary>
        internal int Column { get { return _col; } set { _col = value; } }
		/// <summary>
		/// The address
		/// </summary>
        internal string CellAddress { get { return GetAddress(_row, _col); } }
		/// <summary>
		/// Returns true if the cell's contents are numeric.
		/// </summary>
        public bool IsNumeric { get { return (Value is decimal || Value.GetType().IsPrimitive ); } }
		#region ExcelCell Value
        internal object _value = null;
        /// <summary>
		/// Gets/sets the value of the cell.
		/// </summary>
        public object Value
		{
			get
			{
                return _value;
			}
			set
			{
                SetValueRichText(value);
                if (IsRichText) IsRichText = false;
			}
		}
        internal void SetValueRichText(object value)
        {
            _value = value;
            if (value is string) DataType = "s"; else DataType = "";
            Formula = "";
        }
        /// <summary>
        /// If cell has inline formating. 
        /// </summary>
        public bool IsRichText { get; set; }
        /// <summary>
        /// If the cell is merged with other cells
        /// </summary>
        public bool Merge { get; internal set; }
        /// <summary>
        /// Merge Id
        /// </summary>
        internal int MergeId { get; set; }
		#endregion  

        #region ExcelCell DataType
        /// <summary>
        /// Datatype
        /// TODO: remove
        /// </summary>       
        internal string DataType
        {
            get
            {
                return (_dataType);
            }
            set
            {
                _dataType = value;
            }
        }
        #endregion

		#region ExcelCell Style
        string _styleName=null;
        /// <summary>
		/// Optional named style for the cell
		/// </summary>
		public string StyleName
		{
			get 
            {
                if (_styleName == null)
                {
                    foreach (var ns in _worksheet.Workbook.Styles.NamedStyles)
                    {
                        if (ns.StyleXfId == StyleID)
                        {
                            _styleName = ns.Name;
                            break;
                        }
                    }
                    if (_styleName == null)
                    {
                        _styleName = "Normal";
                    }
                }
                return _styleName;
            }
			set 
            {
                _styleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
                _styleName = value;
            }
		}

		int _styleID=0;
        /// <summary>
		/// The style ID for the cell. Reference to the style collection
		/// </summary>
		public int StyleID
		{
			get
			{
				if(_styleID>0)
                    return _styleID;
                else if (_worksheet._rows != null && _worksheet._rows.ContainsKey(ExcelRow.GetRowID(_worksheet.SheetID, Row)) && _worksheet.Row(Row).StyleID>0)
                {
                    return _worksheet.Row(Row).StyleID;
                }
                else
                {
                    ExcelColumn col = GetColumn(Column);
                    if(col==null)
                    {
                        return 0;
                    }
                    else
                    {
                        return col.StyleID;
                    }
                }
			}
			set 
            {
                _styleID = value;
            }
		}

        private ExcelColumn GetColumn(int col)
        {
            foreach (ExcelColumn column in _worksheet._columns)
            {
                if (col >= column.ColumnMin && col <= column.ColumnMax)
                {
                    return column;
                }
            }
            return null;
        }
        internal int GetCellStyleID()
        {
            return _styleID;
        }
        public ExcelStyle Style
        {
            get
            {
                return _worksheet.Workbook.Styles.GetStyleObject(StyleID, _worksheet.PositionID, CellAddress);
            }
        }
        internal void SetNewStyleName(string Name, int Id)
        {
            _styleID = Id;
            _styleName = Name;

        }
		#endregion

		#region ExcelCell Hyperlink
		/// <summary>
		/// The cells cell's Hyperlink
		/// </summary>
		public Uri Hyperlink
		{
			get
			{				
                return (_hyperlink);
			}
			set
			{
				_hyperlink = value;
                if ((Value == null || Value.ToString() == ""))
                {
                    if (value is ExcelHyperLink)
                    {
                        Value = (value as ExcelHyperLink).Display;
                    }
                    else
                    {
                        Value = _hyperlink.AbsoluteUri;
                    }
                }
			}
		}
        internal string HyperLinkRId
        {
            get;
            set;
        }
		#endregion

		#region ExcelCell Formula
		/// <summary>
		/// The cell's formula.
		/// </summary>
		public string Formula
		{
			get
			{
                if (SharedFormulaID < 0)
                {
                    if (_formula == "")
                    {
                        return (TranslateFromR1C1(_formulaR1C1, Row, Column));
                    }
                    else
                    {
                        return (_formula);
                    }
                }
                else
                {
                    if (_worksheet._sharedFormulas.ContainsKey(SharedFormulaID))
                    {
                        var f = _worksheet._sharedFormulas[SharedFormulaID];
                        if (f.StartRow == Row && f.StartCol == Column)
                        {
                            return f.Formula;
                        }
                        else
                        {
                            return TranslateFromR1C1(TranslateToR1C1(f.Formula, f.StartRow, f.StartCol), Row, Column); 
                        }
                        
                    }
                    else
                    {
                        throw(new Exception("Shared formula reference (SI) is invalid"));
                    }
                }
			}
			set
			{
				_formula = value;
                _formulaR1C1 = "";
                _sharedFormulaID = int.MinValue;
                if (_formula!="" && !_worksheet._formulaCells.ContainsKey(CellID))
                {
                    _worksheet._formulaCells.Add(this);
                }
			}
        }
        /// <summary>
        /// The cell's formula using R1C1 style.
        /// </summary>
        public string FormulaR1C1
        {
            get
            {
                if (SharedFormulaID < 0)
                {
                    if (_formulaR1C1 == "")
                    {
                        return TranslateToR1C1(_formula, Row, Column);
                    }
                    else
                    {
                        return (_formulaR1C1);
                    }
                }
                else
                {
                    if (_worksheet._sharedFormulas.ContainsKey(SharedFormulaID))
                    {
						var f = _worksheet._sharedFormulas[SharedFormulaID];
                        return TranslateToR1C1(f.Formula, f.StartRow, f.StartCol);
                    }
                    else
                    {
                        throw (new Exception("Shared formula reference (SI) is invalid"));
                    }
                }
            }
            set
            {
                // Example cell content for formulas
                // <f>RC1</f>
                // <f>SUM(RC1:RC3)</f>
                // <f>R[-1]C[-2]+R[-1]C[-1]</f>
                _formulaR1C1 = value;
                _formula = "";
                SharedFormulaID = int.MinValue;
                if (!_worksheet._formulaCells.ContainsKey(CellID))
                {
                    _worksheet._formulaCells.Add(this);
                }
            }
        }
        internal int _sharedFormulaID;
        /// <summary>
        /// Id for the shared formula
        /// </summary>
        internal int SharedFormulaID {
            get
            {
                return _sharedFormulaID;
            }
            set
            {
                _sharedFormulaID = value;
                if(_worksheet._formulaCells.ContainsKey(CellID)) _worksheet._formulaCells.Delete(CellID);
            }
        }
        public bool IsArrayFormula { get; internal set; }
		#endregion 		
		/// <summary>
		/// Returns the cell's value as a string.
		/// </summary>
		/// <returns>The cell's value</returns>
		public override string ToString()	{	return Value.ToString();	}
		#endregion 
		#region ExcelCell Private Methods
		#endregion 
        #region IRangeID Members

        ulong IRangeID.RangeID
        {
            get
            {
                return GetCellID(_worksheet.SheetID, Row, Column);
            }
            set
            {
                _col = ((int)(value >> 15) & 0x3FF);
                _row = ((int)(value >> 29));
            }
        }

        #endregion
        internal ExcelCell Clone(ExcelWorksheet added)
        {
            return Clone(added, _row, _col);
        }
        internal ExcelCell Clone(ExcelWorksheet added, int row, int col)
        {
            ExcelCell newCell = new ExcelCell(added, row, col);
            if(_hyperlink!=null) newCell.Hyperlink = Hyperlink;
            newCell._formula = _formula;
            newCell._formulaR1C1 = _formulaR1C1;
            newCell.IsRichText = IsRichText;
            newCell.Merge = Merge;
            newCell._sharedFormulaID = _sharedFormulaID;
            newCell._styleName = _styleName;
            newCell._styleID = _styleID;
            newCell._value = _value;
            return newCell;
        }
    }
}
