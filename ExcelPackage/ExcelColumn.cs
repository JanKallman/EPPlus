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
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/
using System;
using System.Xml;
using OfficeOpenXml.Style;
namespace OfficeOpenXml
{
	/// <summary>
	/// Represents an individual column within the worksheet
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
            _width = _worksheet.defaultColWidth;
        }
		#endregion
        int _columnMin;		
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

                foreach (ExcelColumn c in _worksheet._columns)
                {
                    if (c.ColumnMin > _columnMin && c.ColumnMax <= value && c.ColumnMin!=_columnMin)
                    {
                        throw new Exception(string.Format("ColumnMax can not spann over existing column {0}.",c.ColumnMin));
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
        bool _hidden=false;
        public bool Hidden
		{
			get
			{
                //bool retValue = false;
                //string hidden = _colElement.GetAttribute("hidden", "1");
                //if (hidden == "1") retValue = true;
                //return (retValue);
                return _hidden;
			}
			set
			{
                //if (value)
                //    _colElement.SetAttribute("hidden", "1");
                //else
                //    _colElement.SetAttribute("hidden", "0");
                _hidden = value;
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
        double _width;
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
                _width = value;
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
        public bool Collapsed { get; set; }
        public int OutlineLevel { get; set; }
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
                string letter = ExcelCell.GetColumnLetter(ColumnMin);
                string endLetter = ExcelCell.GetColumnLetter(ColumnMax);
                return _worksheet.Workbook.Styles.GetStyleObject(_styleID, _worksheet.PositionID, letter + ":" + endLetter);
            }
        }
        string _styleName="";
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
                _styleID = _worksheet.Workbook.Styles.GetStyleIdFromName(value);
                _styleName = value;
            }
		}
        int _styleID = 0;
        /// <summary>
		/// Sets the style for the entire column using the style ID.  
		/// </summary>
        public int StyleID
		{
            get
            {
                return _styleID;
            }
            set
            {
                _styleID = value;
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
        internal void Clone(ExcelWorksheet added)
        {
            ExcelColumn newCol = added.Column(ColumnMin);
            newCol.ColumnMax = ColumnMax;
            newCol.BestFit = BestFit;
            newCol.Collapsed = Collapsed;
            newCol.Hidden = Hidden;
            newCol.OutlineLevel = OutlineLevel;
            newCol.PageBreak = PageBreak;
            newCol.Phonetic = Phonetic;
            newCol.StyleName = StyleName;
            newCol.StyleID = StyleID;
            newCol.Width = Width;
        }
    }
}
