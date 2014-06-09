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
using System.Xml;
using OfficeOpenXml.Style;
namespace OfficeOpenXml
{
    /// <summary>
	/// Represents one or more columns within the worksheet
	/// </summary>
	public class ExcelColumn : IRangeID
	{
		private ExcelWorksheet _worksheet;
		private XmlElement _colElement = null;

		#region ExcelColumn Constructor
		/// <summary>
		/// Creates a new instance of the ExcelColumn class.  
		/// For internal use only!
		/// </summary>
		/// <param name="Worksheet"></param>
		/// <param name="col"></param>
		protected internal ExcelColumn(ExcelWorksheet Worksheet, int col)
        {
            _worksheet = Worksheet;
            _columnMin = col;
            _columnMax = col;
            _width = _worksheet.DefaultColWidth;
        }
		#endregion
        internal int _columnMin;		
		/// <summary>
		/// Sets the first column the definition refers to.
		/// </summary>
		public int ColumnMin 
		{
            get { return _columnMin; }
			//set { _columnMin=value; } 
		}

        internal int _columnMax;
        /// <summary>
		/// Sets the last column the definition refers to.
		/// </summary>
        public int ColumnMax 
		{ 
            get { return _columnMax; }
			set 
            {
                if (value < _columnMin && value > ExcelPackage.MaxColumns)
                {
                    throw new Exception("ColumnMax out of range");
                }

                var cse = new CellsStoreEnumerator<object>(_worksheet._values, 0, 0, 0, ExcelPackage.MaxColumns);
                while(cse.Next())
                {
                    var c = cse.Value as ExcelColumn;
                    if (cse.Column > _columnMin && c.ColumnMax <= value && cse.Column!=_columnMin)
                    {
                        throw new Exception(string.Format("ColumnMax can not span over existing column {0}.",c.ColumnMin));
                    }
                }
                _columnMax = value; 
            } 
		}
        /// <summary>
        /// Internal range id for the column
        /// </summary>
        internal ulong ColumnID
        {
            get
            {
                return ExcelColumn.GetColumnID(_worksheet.SheetID, ColumnMin);
            }
        }
		#region ExcelColumn Hidden
		/// <summary>
		/// Allows the column to be hidden in the worksheet
		/// </summary>
        internal bool _hidden=false;
        public bool Hidden
		{
			get
			{
                return _hidden;
			}
			set
			{
                if (_worksheet._package.DoAdjustDrawings)
                {
                    var pos = _worksheet.Drawings.GetDrawingWidths();
                    _hidden = value;
                    _worksheet.Drawings.AdjustWidth(pos);
                }
                else
                {
                    _hidden = value;
                }
			}
		}
		#endregion

		#region ExcelColumn Width
        internal double VisualWidth
        {
            get
            {
                if (_hidden || (Collapsed && OutlineLevel>0))
                {
                    return 0;
                }
                else
                {
                    return _width;
                }
            }
        }
        internal double _width;
        /// <summary>
        /// Sets the width of the column in the worksheet
        /// </summary>
        public double Width
		{
			get
			{
                return _width;
			}
			set	
            {
                if (_worksheet._package.DoAdjustDrawings)
                {
                    var pos = _worksheet.Drawings.GetDrawingWidths();
                    _width = value;
                    _worksheet.Drawings.AdjustWidth(pos);
                }
                else
                {
                    _width = value;
                }

                if (_hidden && value!=0)
                {
                    _hidden = false;
                }
            }
		}
        /// <summary>
        /// If set to true a column automaticlly resize(grow wider) when a user inputs numbers in a cell. 
        /// </summary>
        public bool BestFit
        {
            get;
            set;
        }
        /// <summary>
        /// If the column is collapsed in outline mode
        /// </summary>
        public bool Collapsed { get; set; }
        /// <summary>
        /// Outline level. Zero if no outline
        /// </summary>
        public int OutlineLevel { get; set; }
        /// <summary>
        /// Phonetic
        /// </summary>
        public bool Phonetic { get; set; }
        #endregion

		#region ExcelColumn Style
        /// <summary>
        /// The Style applied to the whole column. Only effects cells with no individual style set. 
        /// Use Range object if you want to set specific styles.
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                string letter = ExcelCellBase.GetColumnLetter(ColumnMin);
                string endLetter = ExcelCellBase.GetColumnLetter(ColumnMax);
                return _worksheet.Workbook.Styles.GetStyleObject(StyleID, _worksheet.PositionID, letter + ":" + endLetter);
            }
        }
        internal string _styleName="";
        /// <summary>
		/// Sets the style for the entire column using a style name.
		/// </summary>
		public string StyleName
		{
            get
            {
                return _styleName;
            }
            set
            {
                StyleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
                _styleName = value;
            }
		}
        //internal int _styleID = 0;
        /// <summary>
        /// Sets the style for the entire column using the style ID.  
        /// </summary>
        public int StyleID
        {
            get
            {
                return _worksheet._styles.GetValue(0, ColumnMin);
            }
            set
            {
                _worksheet._styles.SetValue(0, ColumnMin, value);
            }
        }
        /// <summary>
        /// Adds a manual page break after the column.
        /// </summary>
        public bool PageBreak
        {
            get;
            set;
        }
        public bool Merged
        {
            get
            {
                return _worksheet.MergedCells[ColumnMin, 0] != null;
            }
            set
            {
                _worksheet.MergedCells.Add(new ExcelAddressBase(1, ColumnMin, ExcelPackage.MaxRows, ColumnMax), true);
            }
        }
        #endregion

		/// <summary>
		/// Returns the range of columns covered by the column definition.
		/// </summary>
		/// <returns>A string describing the range of columns covered by the column definition.</returns>
		public override string ToString()
		{
			return string.Format("Column Range: {0} to {1}", _colElement.GetAttribute("min"), _colElement.GetAttribute("min"));
		}
        /// <summary>
        /// Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property.
        /// Note: Cells containing formulas are ignored since EPPlus don't have a calculation engine.
        ///       Wrapped and merged cells are also ignored.
        /// </summary>
        public void AutoFit()
        {
            _worksheet.Cells[1, _columnMin, ExcelPackage.MaxRows, _columnMax].AutoFitColumns();
        }

        /// <summary>
        /// Set the column width from the content.
        /// Note: Cells containing formulas are ignored since EPPlus don't have a calculation engine.
        ///       Wrapped and merged cells are also ignored.
        /// </summary>
        /// <param name="MinimumWidth">Minimum column width</param>
        public void AutoFit(double MinimumWidth)
        {
            _worksheet.Cells[1, _columnMin, ExcelPackage.MaxRows, _columnMax].AutoFitColumns(MinimumWidth);
        }

        /// <summary>
        /// Set the column width from the content.
        /// Note: Cells containing formulas are ignored since EPPlus don't have a calculation engine.
        ///       Wrapped and merged cells are also ignored.
        /// </summary>
        /// <param name="MinimumWidth">Minimum column width</param>
        /// <param name="MaximumWidth">Maximum column width</param>
        public void AutoFit(double MinimumWidth, double MaximumWidth)
        {
            _worksheet.Cells[1, _columnMin, ExcelPackage.MaxRows, _columnMax].AutoFitColumns(MinimumWidth, MaximumWidth);
        }

        /// <summary>
        /// Get the internal RangeID
        /// </summary>
        /// <param name="sheetID">Sheet no</param>
        /// <param name="column">Column</param>
        /// <returns></returns>
        internal static ulong GetColumnID(int sheetID, int column)
        {
            return ((ulong)sheetID) + (((ulong)column) << 15);
        }

        #region IRangeID Members

        ulong IRangeID.RangeID
        {
            get
            {
                return ColumnID;
            }
            set
            {
                int prevColMin = _columnMin;
                _columnMin = ((int)(value >> 15) & 0x3FF);
                _columnMax += prevColMin - ColumnMin;
                //Todo:More Validation
                if (_columnMax > ExcelPackage.MaxColumns) _columnMax = ExcelPackage.MaxColumns;
            }
        }

        #endregion

        /// <summary>
        /// Copies the current column to a new worksheet
        /// </summary>
        /// <param name="added">The worksheet where the copy will be created</param>
        internal ExcelColumn Clone(ExcelWorksheet added)
        {
            return Clone(added, ColumnMin);
        }
        internal ExcelColumn Clone(ExcelWorksheet added, int col)
        {
            ExcelColumn newCol = added.Column(col);
                newCol.ColumnMax = ColumnMax;
                newCol.BestFit = BestFit;
                newCol.Collapsed = Collapsed;
                newCol.OutlineLevel = OutlineLevel;
                newCol.PageBreak = PageBreak;
                newCol.Phonetic = Phonetic;
                newCol._styleName = _styleName;
                newCol.StyleID = StyleID;
                newCol.Width = Width;
                newCol.Hidden = Hidden;
                return newCol;
        }
    }
}
