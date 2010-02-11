/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * EPPlus is a fork of the ExcelPackage project http://excelpackage.codeplex.com/
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
 *******************************************************************************
 * Jan Källman		Added		10-SEP-2009
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
		#region Cell Private Properties
		private ExcelWorksheet _xlWorksheet;
		private int _row;
		private int _col;
		internal string _formula="";
		private Uri _hyperlink=null;
        static CultureInfo _ci=new CultureInfo("en-US");
        #endregion
		#region ExcelCell Constructor
		/// <summary>
		/// Creates a new instance of ExcelCell class. For internal use only!
		/// </summary>
		/// <param name="xlWorksheet">A reference to the parent worksheet</param>
		/// <param name="row">The row number in the parent worksheet</param>
		/// <param name="col">The column number in the parent worksheet</param>
		protected internal ExcelCell(ExcelWorksheet xlWorksheet, int row, int col)
		{
			if (row < 1 || col < 1)
				throw new Exception("ExcelCell Constructor: Negative row and column numbers are not allowed");
            if (row > ExcelPackage.MaxRows || col > ExcelPackage.MaxColumns)
                throw new Exception("ExcelCell Constructor: row or column numbers are out of range");
            if (xlWorksheet == null)
				throw new Exception("ExcelCell Constructor: xlWorksheet must be set to a valid reference");

			_row = row;
			_col = col;
            _xlWorksheet = xlWorksheet;
            if (col < xlWorksheet._minCol) xlWorksheet._minCol = col;
            if (col > xlWorksheet._maxCol) xlWorksheet._maxCol = col;
            SharedFormulaID = int.MinValue;
            IsRichText = false;

		}
        protected internal ExcelCell(ExcelWorksheet xlWorksheet, string cellAddress)
        {
            _xlWorksheet = xlWorksheet;
            GetRowCol(cellAddress, out _row, out _col);
            if (_col < xlWorksheet._minCol) xlWorksheet._minCol = _col;
            if (_col > xlWorksheet._maxCol) xlWorksheet._maxCol = _col;
            SharedFormulaID = int.MinValue;
            IsRichText = false;
        }
		#endregion  // END Cell Constructors
        internal ulong CellID
        {
            get
            {
                return GetCellID(_xlWorksheet.SheetID, Row, Column);
            }
        }
        #region ExcelCell Public Properties

		/// <summary>
		/// Read-only reference to the cell's row number
		/// </summary>
        public int Row { get { return _row; } internal set { _row = value; } }
		/// <summary>
		/// Read-only reference to the cell's column number
		/// </summary>
        public int Column { get { return _col; } internal set { _row = value; } }
		/// <summary>
		/// Returns the current cell address in the standard Excel format (e.g. 'E5')
		/// </summary>
		public string CellAddress { get { return GetAddress(_row, _col); } }
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
				_value = value;
                if (value is string) DataType = "s"; else DataType = "";
                Formula = "";
			}
		}
        /// <summary>
        /// If cell has inline formating. 
        /// Use XML format specified in the OPEN XML Documentation in the value property
        /// </summary>
        public bool IsRichText { get; set; }
        /// <summary>
        /// If the cell is merged with other cells
        /// </summary>
        public bool Merge { get; internal set; }
		#endregion  

		#region ExcelCell DataType
        string _dataType="";
        /// <summary>
		/// Gets/sets the cell's data type.  
		/// Not currently implemented correctly!
		/// </summary>       
        public string DataType
		{
			// TODO: complete DataType
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
        string _styleName="Normal";
        /// <summary>
		/// Allows you to set the cell's style using a named style
		/// </summary>
		public string StyleName
		{
			get 
            {
                return _styleName;
            }
			set 
            {
                _styleID = _xlWorksheet.Workbook.Styles.GetStyleIdFromName(value);
                _styleName = value;
            }
		}

		int _styleID=0;
        /// <summary>
		/// Allows you to set the cell's style using the number of the style.
		/// Useful when coping styles from one cell to another.
		/// </summary>
		public int StyleID
		{
			get
			{
				if(_styleID>0)
                    return _styleID;
                else if (_xlWorksheet._rows != null && _xlWorksheet.Row(Row).StyleID > 0)
                {
                    return _xlWorksheet.Row(Row).StyleID;
                }
                else
                {
                    return _xlWorksheet.Column(Column).StyleID;                    
                }

			}
			set 
            {
                _styleID = value;
            }
		}
        internal int GetCellStyleID()
        {
            return _styleID;
        }
        public ExcelStyle Style
        {
            get
            {
                return _xlWorksheet.Workbook.Styles.GetStyleObject(StyleID, _xlWorksheet.PositionID, CellAddress);
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
		/// Allows you to set/get the cell's Hyperlink
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
		/// Provides read/write access to the cell's formula.
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
                    if (_xlWorksheet._sharedFormulas.ContainsKey(SharedFormulaID))
                    {
                        return TranslateFromR1C1(_xlWorksheet._sharedFormulas[SharedFormulaID].Formula, Row, Column);
                    }
                    else
                    {
                        throw(new Exception("Shared formula reference (SI) is invalid"));
                    }
                }
			}
			set
			{
				// Example cell content for formulas
				// <f>D7</f>
				// <f>SUM(D6:D8)</f>
				// <f>F6+F7+F8</f>
				_formula = value;
                _formulaR1C1 = "";
                SharedFormulaID = int.MinValue;
                if (_formula!="" && !_xlWorksheet._formulaCells.ContainsKey(CellID))
                {
                    _xlWorksheet._formulaCells.Add(this);
                }
			}
        }
        string _formulaR1C1="";
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
                    if (_xlWorksheet._sharedFormulas.ContainsKey(SharedFormulaID))
                    {
                        return TranslateToR1C1(_xlWorksheet._sharedFormulas[SharedFormulaID].Formula, Row, Column);
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
                if (!_xlWorksheet._formulaCells.ContainsKey(CellID))
                {
                    _xlWorksheet._formulaCells.Add(this);
                }
            }
        }
        /// <summary>
        /// Id for the shared formula
        /// </summary>
        public int SharedFormulaID { get; set; }

		#region ExcelCell Comment
		/// <summary>
		/// Returns the comment as a string
		/// </summary>
		public string Comment
		{
			// TODO: implement get which will obtain the text of the comment from the comment1.xml file
			get
			{
				throw new Exception("Function not yet implemented!");
			}
			// TODO: implement set which will add comments to the worksheet
			// this will require you to add entries to the Drawing.vml file to get this to work! 
		}
		#endregion 

		// TODO: conditional formatting

		#endregion  // END Cell Public Properties
		
		/// <summary>
		/// Returns the cell's value as a string.
		/// </summary>
		/// <returns>The cell's value</returns>
		public override string ToString()	{	return Value.ToString();	}

		#endregion  // END Cell Public Methods
		#region ExcelCell Private Methods

		#region IsNumericValue
		/// <summary>
		/// Returns true if the string contains a numeric value
		/// </summary>
		/// <param name="Value"></param>
		/// <returns></returns>
		public static bool IsNumericValue(string Value)
		{
			Regex objNotIntPattern = new Regex("[^0-9,.-]");
			Regex objIntPattern = new Regex("^-[0-9,.]+$|^[0-9,.]+$");

			return !objNotIntPattern.IsMatch(Value) &&
							objIntPattern.IsMatch(Value);
		}
		#endregion
		#endregion // END Cell Private Methods
        #region IRangeID Members

        ulong IRangeID.RangeID
        {
            get
            {
                return GetCellID(_xlWorksheet.SheetID, Row, Column);
            }
            set
            {
                //_sheet = (int)(cellID % 0x8000);
                _col = ((int)(value >> 15) & 0x3FF);
                _row = ((int)(value >> 29));
            }
        }

        #endregion
    }
}
