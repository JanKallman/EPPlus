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
 * Jan Källman		    Initial Release		        2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/

using System;
using System.Xml;
using OfficeOpenXml.Style;
namespace OfficeOpenXml
{
	/// <summary>
	/// Represents an individual row in the spreadsheet.
	/// </summary>
	public class ExcelRow : IRangeID
	{
		private ExcelWorksheet _worksheet;
		private XmlElement _rowElement = null;
        /// <summary>
        /// Internal RowID.
        /// </summary>
        public ulong RowID 
        {
            get
            {
                return GetRowID(_worksheet.SheetID, Row);
            }
        }
		#region ExcelRow Constructor
		/// <summary>
		/// Creates a new instance of the ExcelRow class. 
		/// For internal use only!
		/// </summary>
		/// <param name="Worksheet">The parent worksheet</param>
		/// <param name="row">The row number</param>
		internal ExcelRow(ExcelWorksheet Worksheet, int row)
		{
			_worksheet = Worksheet;
            Row = row;
		}
		#endregion

		/// <summary>
		/// Provides access to the node representing the row.
		/// </summary>
		internal XmlNode Node { get { return (_rowElement); } }

		#region ExcelRow Hidden
        bool _hidden = false;
        /// <summary>
		/// Allows the row to be hidden in the worksheet
		/// </summary>
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
                    _worksheet.Drawings.AdjustHeight(pos);
                }
                else
                {
                    _hidden = value;
                }
            }
        }        
		#endregion

		#region ExcelRow Height
        double _height=-1;  //Set to default height
        /// <summary>
		/// Sets the height of the row
		/// </summary>
		public double Height
        {
			get
			{
                if (_height == -1)
                {
                    return _worksheet.DefaultRowHeight;
                }
                else
                {
                    return _height;
                }
                //}
			}
			set	
            {
                if (_worksheet._package.DoAdjustDrawings)
                {
                    var pos = _worksheet.Drawings.GetDrawingWidths();
                    _height = value;
                    _worksheet.Drawings.AdjustHeight(pos);
                }
                else
                {
                    _height = value;
                }

                if (Hidden && value != 0)
                {
                    Hidden = false;
                }
                CustomHeight = (value != _worksheet.DefaultRowHeight);
            }
        }
        /// <summary>
        /// Set to true if You don't want the row to Autosize
        /// </summary>
        public bool CustomHeight { get; set; }
		#endregion

        internal string _styleName = "";
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
                _styleId = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
                _styleName = value;
            }
        }        
        internal int _styleId = 0;
		/// <summary>
		/// Sets the style for the entire row using the style ID.  
		/// </summary>
        public int StyleID
		{
			get
			{
				return _styleId; 
			}
			set	
			{
                _styleId = value;
			}
		}

        /// <summary>
        /// Rownumber
        /// </summary>
        public int Row
        {
            get;
            set;
        }
        /// <summary>
        /// If outline level is set this tells that the row is collapsed
        /// </summary>
        public bool Collapsed
        {
            get;
            set;
        }
        /// <summary>
        /// Outline level.
        /// </summary>
        public int OutlineLevel
        {
            get;
            set;
        }        
        /// <summary>
        /// Show phonetic Information
        /// </summary>
        public bool Phonetic 
        {
            get;
            set;
        }
        /// <summary>
        /// The Style applied to the whole row. Only effekt cells with no individual style set. 
        /// Use ExcelRange object if you want to set specific styles.
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return _worksheet.Workbook.Styles.GetStyleObject(StyleID,_worksheet.PositionID ,Row.ToString()+":"+Row.ToString());                
            }
        }
        /// <summary>
        /// Adds a manual page break after the row.
        /// </summary>
        public bool PageBreak
        {
            get;
            set;
        }
        internal static ulong GetRowID(int sheetID, int row)
        {
            return ((ulong)sheetID) + (((ulong)row) << 29);

        }

        #region IRangeID Members

        ulong IRangeID.RangeID
        {
            get
            {
                return RowID; 
            }
            set
            {
                Row = ((int)(value >> 29));
            }
        }

        #endregion
        /// <summary>
        /// Copies the current row to a new worksheet
        /// </summary>
        /// <param name="added">The worksheet where the copy will be created</param>
        internal void Clone(ExcelWorksheet added)
        {
            ExcelRow newRow = added.Row(Row);
            newRow.Collapsed = Collapsed;
            newRow.Height = Height;
            newRow.Hidden = Hidden;
            newRow.OutlineLevel = OutlineLevel;
            newRow.PageBreak = PageBreak;
            newRow.Phonetic = Phonetic;
            newRow.StyleName = StyleName;
            newRow.StyleID = StyleID;
        }
    }
}
