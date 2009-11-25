using System;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using OfficeOpenXml.Style;
using System.Text.RegularExpressions;

namespace OfficeOpenXml
{
    public class ExcelCellNew
    {
		#region Cell Private Properties
		private ExcelWorksheet _xlWorksheet;
		private int _row;
		private int _col;
		private string _valueRef;
		private string _formula="";
		//private string _dataType;
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
		protected internal ExcelCellNew(ExcelWorksheet xlWorksheet, int row, int col)
		{
			if (row < 1 || col < 1)
				throw new Exception("ExcelCell Constructor: Negative row and column numbers are not allowed");
			if (xlWorksheet == null)
				throw new Exception("ExcelCell Constructor: xlWorksheet must be set to a valid reference");

			_xlWorksheet = xlWorksheet;
			_row = row;
			_col = col;
            SharedFormulaID = int.MinValue;
		}
        protected internal ExcelCellNew(ExcelWorksheet xlWorksheet, string cellAddress)
        {
            _xlWorksheet = xlWorksheet;
            _row = GetRowNumber(cellAddress);
            _col = GetColumnNumber(cellAddress);
            SharedFormulaID = int.MinValue;
        }
		#endregion  // END Cell Constructors
        static internal void SplitCellID(long cellID, out int sheet, out int row, out int col)
        {
            sheet=(int)(cellID % 1024);
            col = (int)((cellID % (16777216- (sheet))) / 1024);
            row = (int)((cellID - (cellID % 16777216L)) / 16777216L);
        }
        internal static ulong GetCellID(int SheetID, int row, int col)
        {
            return (ulong)(SheetID) + (ulong)(col) * 1024 + ((ulong)(row) * 16777216);
        }
        #region ExcelCell Public Properties

		/// <summary>
		/// Read-only reference to the cell's XmlNode (for internal use only)
		/// </summary>
        public int Row { get { return _row; } internal set { _row = value; } }
		/// <summary>
		/// Read-only reference to the cell's column number
		/// </summary>
        public int Column { get { return _col; } internal set { _row = value; } }
		/// <summary>
		/// Returns the current cell address in the standard Excel format (e.g. 'E5')
		/// </summary>
		public string CellAddress { get { return GetCellAddress(_row, _col); } }
		/// <summary>
		/// Returns true if the cell's contents are numeric.
		/// </summary>
		public bool IsNumeric { get { return (Value is decimal); } }

		#region ExcelCell Value
        object _value = null;
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
				return _styleID;
			}
			set 
            {
                _styleID = value;
            }
		}
        internal ulong CellID
        {
            get
            {
                return (ulong)(_xlWorksheet.SheetID) + (ulong)(Column) * 1024 + ((ulong)(Row) * 16777216);
            }
        }
        public ExcelStyle Style
        {
            get
            {
                return _xlWorksheet.Workbook.Styles.GetStyleObject(StyleID, _xlWorksheet.SheetID, CellAddress);                
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
                Value = _hyperlink.AbsoluteUri;
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
                    return (_formula);
                }
                else
                {
                    if (_xlWorksheet._sharedFormulas.ContainsKey(SharedFormulaID))
                    {
                        return _xlWorksheet._sharedFormulas[SharedFormulaID].Formula;
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
                SharedFormulaID = int.MinValue;
                // insert the formula into the cell
			}
		}
        internal int SharedFormulaID { get; set; }
		#endregion

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

		#region ExcelCell Public Methods
		/// <summary>
		/// Removes the XmlNode that holds the cell's value.
		/// Useful when the cell contains a formula as this will force Excel to re-calculate the cell's content.
		/// </summary>
		public void RemoveValue()
		{
            _value = null;
        }
		
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
		
		#region SharedString methods
        //private int SetSharedString(string Value)
        //{
        //    //  Assume the string won't be found (assign it an impossible index):
        //    int index = -1;

        //    //  Check to see if the string already exists. If so, retrieve its index.
        //    //  This search is case-sensitive, but Excel stores differently cased
        //    //  strings separately within the string file.
			
        //    XmlNode stringNode = _xlWorksheet.xlPackage.Workbook.SharedStringsXml.SelectSingleNode(string.Format("//d:si[d:t='{0}']", Value.Replace("'","&apos;")), _xlWorksheet.NameSpaceManager);
        //    if (stringNode == null)
        //    {
        //        //  You didn't find the string in the table, so add it now.
        //        stringNode = _xlWorksheet.xlPackage.Workbook.SharedStringsXml.CreateElement("si", ExcelPackage.schemaMain);
        //        XmlElement textNode = _xlWorksheet.xlPackage.Workbook.SharedStringsXml.CreateElement("t", ExcelPackage.schemaMain);
        //        textNode.InnerText = Value;
        //        stringNode.AppendChild(textNode);
        //        _xlWorksheet.xlPackage.Workbook.SharedStringsXml.DocumentElement.AppendChild(stringNode);
        //    }

        //    if (stringNode != null)
        //    {
        //        //  Retrieve the index of the selected node.
        //        //  To do that, count the number of preceding
        //        //  nodes by retrieving a reference to those nodes.
        //        XmlNodeList nodes = stringNode.SelectNodes("preceding-sibling::d:si", _xlWorksheet.NameSpaceManager);
        //        index = nodes.Count;
        //    }
        //    return (index);
        //}
        //private string GetSharedString(int stringID)
        //{
        //    string retValue = null;
        //    XmlNodeList stringNodes = _xlWorksheet.xlPackage.Workbook.SharedStringsXml.SelectNodes(string.Format("//d:si", stringID), _xlWorksheet.NameSpaceManager);
        //    XmlNode stringNode = stringNodes[stringID];
        //    if (stringNode != null)
        //        retValue = stringNode.InnerText;
        //    return (retValue);
        //}
		#endregion

        //#region AddFormulaNode
        ///// <summary>
        ///// Adds a new formula node to the cell in the correct location
        ///// </summary>
        ///// <returns></returns>
        //protected internal XmlElement AddFormulaElement()
        //{
        //    XmlElement formulaElement = _cellElement.OwnerDocument.CreateElement("f", ExcelPackage.schemaMain);
        //    // find the right location for insersion
        //    XmlNode valueNode = _cellElement.SelectSingleNode("./d:v", _xlWorksheet.NameSpaceManager);
        //    if (valueNode == null)
        //        _cellElement.AppendChild(formulaElement);
        //    else
        //        _cellElement.InsertBefore(formulaElement, valueNode);
        //    return formulaElement;
        //}
        //#endregion

        //#region GetOrCreateCellElement
        //private XmlElement GetOrCreateCellElement(ExcelWorksheet xlWorksheet, int row, int col)
        //{
        //    XmlElement cellNode = null;
        //    // this will create the row if it does not already exist
        //    XmlNode rowNode = xlWorksheet.Row(row).Node;
        //    if (rowNode != null)
        //    {
        //        cellNode = (XmlElement) rowNode.SelectSingleNode(string.Format("./d:c[@" + ExcelWorksheet.tempColumnNumberTag + "='{0}']", col), _xlWorksheet.NameSpaceManager);
        //        if (cellNode == null)
        //        {
        //            //  Didn't find the cell so create the cell element
        //            cellNode = xlWorksheet.WorksheetXml.CreateElement("c", ExcelPackage.schemaMain);
        //            cellNode.SetAttribute(ExcelWorksheet.tempColumnNumberTag, col.ToString());

        //            // You must insert the new cell at the correct location.
        //            // Loop through the children, looking for the first cell that is
        //            // beyond the cell you're trying to insert. Insert before that cell.
        //            XmlNode biggerNode = null;
        //            XmlNodeList cellNodes = rowNode.SelectNodes("./d:c", _xlWorksheet.NameSpaceManager);
        //            if (cellNodes != null)
        //            {
        //                foreach (XmlNode node in cellNodes)
        //                {
        //                    XmlNode colNode = node.Attributes[ExcelWorksheet.tempColumnNumberTag];
        //                    if (colNode != null)
        //                    {
        //                        int colFound = Convert.ToInt32(colNode.Value);
        //                        if (colFound > col)
        //                        {
        //                            biggerNode = node;
        //                            break;
        //                        }
        //                    }
        //                }
        //            }
        //            if (biggerNode == null)
        //            {
        //                rowNode.AppendChild(cellNode);
        //            }
        //            else
        //            {
        //                rowNode.InsertBefore(cellNode, biggerNode);
        //            }
        //        }
        //    }
        //    return (cellNode);
        //}
        //#endregion

		#endregion // END Cell Private Methods

		#region ExcelCell Static Cell Address Manipulation Routines

		#region GetColumnLetter
		/// <summary>
		/// Returns the character representation of the numbered column
		/// </summary>
		/// <param name="iColumnNumber">The number of the column</param>
		/// <returns>The letter representing the column</returns>
		protected internal static string GetColumnLetter(int iColumnNumber)
		{
			int iMainLetterUnicode;
			char iMainLetterChar;

			// TODO: we need to cater for columns larger than ZZ
			if (iColumnNumber > 26)
			{
				int iFirstLetterUnicode = 0;  // default
				int iFirstLetter = Convert.ToInt32(iColumnNumber / 26);
				char iFirstLetterChar;
				if (Convert.ToDouble(iFirstLetter) == (Convert.ToDouble(iColumnNumber) / 26))
				{
					iFirstLetterUnicode = iFirstLetter - 1 + 64;
					iMainLetterChar = 'Z';
				}
				else
				{
					iFirstLetterUnicode = iFirstLetter + 64;
					iMainLetterUnicode = (iColumnNumber - (iFirstLetter * 26)) + 64;
					iMainLetterChar = (char)iMainLetterUnicode;
				}
				iFirstLetterChar = (char)iFirstLetterUnicode;

				return (iFirstLetterChar.ToString() + iMainLetterChar.ToString());
			}
			// if we get here we only have a single letter to return
			iMainLetterUnicode = 64 + iColumnNumber;
			iMainLetterChar = (char)iMainLetterUnicode;
			return (iMainLetterChar.ToString());
		}
		#endregion

		#region GetColumnNumber
		/// <summary>
		/// Returns the column number from the cellAddress
		/// e.g. D5 would return 5
		/// </summary>
		/// <param name="cellAddress">An Excel format cell addresss (e.g. D5)</param>
		/// <returns>The column number</returns>
		public static int GetColumnNumber(string cellAddress)
		{
			// find out position where characters stop and numbers begin
			int iColumnNumber = 0;
			int iPos = 0;
			bool found = false;
			foreach (char chr in cellAddress.ToCharArray())
			{
				iPos++;
				if (char.IsNumber(chr))
				{
					found = true;
					break;
				}
			}

			if (found)
			{
				string AlphaPart = cellAddress.Substring(0, cellAddress.Length - (cellAddress.Length + 1 - iPos));

				int length = AlphaPart.Length;
				int count = 0;
				foreach (char chr in AlphaPart.ToCharArray())
				{
					count++;
					int chrValue = ((int)chr - 64);
					switch (length)
					{
						case 1:
							iColumnNumber = chrValue;
							break;
						case 2:
							if (count == 1)
								iColumnNumber += (chrValue * 26);
							else
								iColumnNumber += chrValue;
							break;
						case 3:
							if (count == 1)
								iColumnNumber += (chrValue * 26 * 26);
							if (count == 2)
								iColumnNumber += (chrValue * 26);
							else
								iColumnNumber += chrValue;
							break;
						case 4:
							if (count == 1)
								iColumnNumber += (chrValue * 26 * 26 * 26);
							if (count == 2)
								iColumnNumber += (chrValue * 26 * 26);
							if (count == 3)
								iColumnNumber += (chrValue * 26);
							else
								iColumnNumber += chrValue;
							break;
					}
				}
			}
			return (iColumnNumber);
		}
		#endregion

		#region GetRowNumber
		/// <summary>
		/// Returns the row number from the cellAddress
		/// e.g. D5 would return 5
		/// </summary>
		/// <param name="cellAddress">An Excel format cell addresss (e.g. D5)</param>
		/// <returns>The row number</returns>
		public static int GetRowNumber(string cellAddress)
		{
			// find out position where characters stop and numbers begin
			int iPos = 0;
			bool found = false;
			foreach (char chr in cellAddress.ToCharArray())
			{
				iPos++;
				if (char.IsNumber(chr))
				{
					found = true;
					break;
				}
			}
			if (found)
			{
				string NumberPart = cellAddress.Substring(iPos - 1, cellAddress.Length - (iPos - 1));
				if (IsNumericValue(NumberPart))
					return (int.Parse(NumberPart));
			}
			return (0);
		}
		#endregion

		#region GetCellAddress
		/// <summary>
		/// Returns the AlphaNumeric representation that Excel expects for a Cell Address
		/// </summary>
		/// <param name="iRow">The number of the row</param>
		/// <param name="iColumn">The number of the column in the worksheet</param>
		/// <returns>The cell address in the format A1</returns>
		public static string GetCellAddress(int iRow, int iColumn)
		{
			return (GetColumnLetter(iColumn) + iRow.ToString());
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
			int row = GetRowNumber(cellAddress);
			int col = GetColumnNumber(cellAddress);

			if (GetCellAddress(row, col) == cellAddress)
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
				int row = GetRowNumber(cellAddress);
				int col = GetColumnNumber(cellAddress);
				string newCellAddress = "";
				if (ExcelCellNew.GetCellAddress(row, col) == cellAddress)   // this checks if the cellAddress is valid
				{
					// we have a valid cell address so change its value (if necessary)
					if (row >= afterRow)
						row += rowIncrement;
					if (col >= afterColumn)
						col += colIncrement;
					newCellAddress = GetCellAddress(row, col);
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

		#endregion // END CellAddress Manipulation Routines

    }
}
