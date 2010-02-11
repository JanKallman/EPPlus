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

/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * ExcelPackage provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/ExcelPackage for details.
 * 
 * Copyright 2007 © Dr John Tunnicliffe 
 * mailto:dr.john.tunnicliffe@btinternet.com
 * All rights reserved.
 * 
 * ExcelPackage is an Open Source project provided under the 
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
 */

/*
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * John Tunnicliffe		Initial Release		01-Jan-2007
 * ******************************************************************************
 */
using System;
using System.Xml;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Configuration;
using OfficeOpenXml.Drawing;
using System.Diagnostics;
using OfficeOpenXml.Style;
using System.Globalization;
using System.Text;
using System.Security;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style.XmlAccess;
namespace OfficeOpenXml
{
    /// <summary>
	/// Represents an Excel worksheet and provides access to its properties and methods
	/// </summary>
	public class ExcelWorksheet : XmlHelper
	{
        internal class Formulas
        {
            internal int Index { get; set; }
            internal string Address { get; set; }
            public string Formula { get; set; }
            public int StartRow { get; set; }
            public int StartCol { get; set; }

        }
        public class MergeCellsCollection<T> : IEnumerable<T>
        {
            private List<T> _list = new List<T>();
            internal MergeCellsCollection()
            {

            }
            internal List<T> List { get {return _list;} }
            public T this[int Index]
            {
                get
                {
                    return _list[Index];
                }
            }
            public int Count
            {
                get
                {
                    return _list.Count;
                }
            }



            #region IEnumerable<T> Members

            public IEnumerator<T> GetEnumerator()
            {
                return _list.GetEnumerator();
            }

            #endregion

            #region IEnumerable Members

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return _list.GetEnumerator();
            }

            #endregion
        }
        internal RangeCollection _cells;
        internal RangeCollection _columns;
        internal RangeCollection _rows;

        internal Dictionary<int, Formulas> _sharedFormulas = new Dictionary<int, Formulas>();
        internal RangeCollection _formulaCells;
        internal static CultureInfo _ci=new CultureInfo("en-US");
        internal int _minCol = ExcelPackage.MaxColumns;
        internal int _maxCol = 0;
        /// <summary>
		/// Temporary tag for all column numbers in the worksheet XML
		/// For internal use only!
		/// </summary>
		protected internal const string tempColumnNumberTag = "colNumber"; 
		/// <summary>
		/// Reference to the parent package
		/// For internal use only!
		/// </summary>
        #region Worksheet Private Properties
        protected internal ExcelPackage xlPackage;
		private Uri _worksheetUri;
		private string _name;
		private int _sheetID;
        private int _positionID;
        private bool _hidden;
		private string _relationshipID;
		private XmlDocument _worksheetXml;
		private ExcelWorksheetView _sheetView;
		private ExcelHeaderFooter _headerFooter;
        #endregion
        #region ExcelWorksheet Constructor
        /// <summary>
        /// A worksheet
        /// </summary>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="excelPackage">Package</param>
        /// <param name="uriWorksheet">URI</param>
        /// <param name="sheetName">Name of the sheet</param>
        /// <param name="SheetID">Sheet id</param>
        /// <param name="PositionID">Position</param>
        /// <param name="Hide">Hidden</param>
        public ExcelWorksheet(XmlNamespaceManager ns, ExcelPackage excelPackage, string relID, 
                              Uri uriWorksheet, string sheetName, int SheetID, int PositionID,
                              bool Hide) :
            base(ns, null)
        {
            SchemaNodeOrder = new string[] { "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "protectedRanges", "customSheetViews", "hyperlinks", "pageMargins", "pageSetup", "drawing" };
            xlPackage = excelPackage;
            _relationshipID = RelationshipID;
            _worksheetUri = uriWorksheet;
            _name = sheetName;
            _sheetID = SheetID;
            _positionID = PositionID;
            Hidden = Hide;
            CreateXml();
            TopNode = _worksheetXml.DocumentElement;
        }
        #endregion

		#region Worksheet Public Properties
		/// <summary>
		/// Read-only: the Uri to the worksheet within the package
		/// </summary>
		protected internal Uri WorksheetUri { get { return (_worksheetUri); } }
		/// <summary>
		/// Read-only: a reference to the PackagePart for the worksheet within the package
		/// </summary>
		protected internal PackagePart Part { get { return (xlPackage.Package.GetPart(WorksheetUri)); } }
		/// <summary>
		/// Read-only: the ID for the worksheet's relationship with the workbook in the package
		/// </summary>
		protected internal string RelationshipID { get { return (_relationshipID); } }
		/// <summary>
		/// The unique identifier for the worksheet.  Note that these can be random, so not
		/// too useful in code!
		/// </summary>
		protected internal int SheetID { get { return (_sheetID); } }
        protected internal int PositionID { get { return (_positionID); } }
		/// <summary>
		/// Returns a ExcelWorksheetView object that allows you to
		/// set the view state properties of the worksheet
		/// </summary>
		public ExcelWorksheetView View
		{
			get
			{
				if (_sheetView == null)
				{
                    XmlNode node = _worksheetXml.SelectSingleNode("//d:sheetView", NameSpaceManager);
                    if (node == null)
                    {
                        CreateNode("d:sheetViews/d:sheetView"); //this one shouls always exist. but check anyway
                    }
                    _sheetView = new ExcelWorksheetView(NameSpaceManager, node, this);
				}
				return (_sheetView);
			}
		}

		#region Name // Worksheet Name
		/// <summary>
		/// The worksheet's name as it appears on the tab
		/// </summary>
		public string Name
		{
			get { return (_name); }
			set
			{
				XmlNode sheetNode = xlPackage.Workbook.WorkbookXml.SelectSingleNode(string.Format("//d:sheet[@sheetId={0}]", _sheetID), NameSpaceManager);
				if (sheetNode != null)
				{
					XmlAttribute nameAttr = (XmlAttribute)sheetNode.Attributes.GetNamedItem("name");
					if (nameAttr != null)
					{
						nameAttr.Value = value;
					}
				}
				_name = value;
			}
		}
		#endregion // END Worksheet Name

		#region Hidden
		/// <summary>
		/// Indicates if the worksheet is hidden in the workbook
		/// </summary>
		public bool Hidden
		{
			get { return (_hidden); }
			set
			{
				XmlNode sheetNode = xlPackage.Workbook.WorkbookXml.SelectSingleNode(string.Format("//d:sheet[@sheetId={0}]", _sheetID), NameSpaceManager);
				if (sheetNode != null)
				{
					XmlAttribute nameAttr = (XmlAttribute)sheetNode.Attributes.GetNamedItem("hidden");
					if (nameAttr != null)
					{
						nameAttr.Value = value.ToString();
					}
				}
				_hidden = value;
			}
		}
		#endregion

		#region defaultRowHeight
        double _defaultRowHeight = double.NaN;
        /// <summary>
		/// Allows you to get/set the default height of all rows in the worksheet
		/// </summary>
        public double defaultRowHeight
		{
			get 
			{
				if(double.IsNaN(_defaultRowHeight))
                {
                    XmlElement sheetFormat = (XmlElement) WorksheetXml.SelectSingleNode("//d:sheetFormatPr", NameSpaceManager);
                    if (sheetFormat == null)
                    {
                        _defaultRowHeight = 15; // Excel's default height
                    }
                    else
                    {
                        string ret = sheetFormat.GetAttribute("defaultRowHeight");
                        _defaultRowHeight = double.Parse(ret, _ci);
                    }
                }
				return _defaultRowHeight;
			}
			set
			{
                _defaultRowHeight = value;
                XmlElement sheetFormat = (XmlElement)WorksheetXml.SelectSingleNode("//d:sheetFormatPr", NameSpaceManager);
				if (sheetFormat == null)
				{
					// create the node as it does not exist
					sheetFormat = WorksheetXml.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
					// find location to insert new element
					XmlNode sheetViews = WorksheetXml.SelectSingleNode("//d:sheetViews", NameSpaceManager);
					// insert the new node
					WorksheetXml.DocumentElement.InsertAfter(sheetFormat, sheetViews);
				}
				sheetFormat.SetAttribute("defaultRowHeight", value.ToString());
			}
		}
		#endregion
        #region defaultColWidth
        /// <summary>
        /// Allows you to get/set the default width of all rows in the worksheet
        /// </summary>
        public double defaultColWidth
        {
            get
            {
                double retValue = 9.140625 ; // Excel's default height
                XmlElement sheetFormat = (XmlElement)WorksheetXml.SelectSingleNode("//d:sheetFormatPr", NameSpaceManager);
                if (sheetFormat != null)
                {
                    string ret = sheetFormat.GetAttribute("defaultColWidth");
                    if (ret != "")
                        retValue = double.Parse(ret, _ci);
                }
                return retValue;
            }
            set
            {
                XmlElement sheetFormat = (XmlElement)WorksheetXml.SelectSingleNode("//d:sheetFormatPr", NameSpaceManager);
                if (sheetFormat == null)
                {
                    // create the node as it does not exist
                    sheetFormat = WorksheetXml.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
                    // find location to insert new element
                    XmlNode sheetViews = WorksheetXml.SelectSingleNode("//d:sheetViews", NameSpaceManager);
                    // insert the new node
                    WorksheetXml.DocumentElement.InsertAfter(sheetFormat, sheetViews);
                }
                sheetFormat.SetAttribute("defaultColWidth", value.ToString());
            }
        }
        #endregion
        /** <outlinePr applyStyles="1" summaryBelow="0" summaryRight="0" /> **/
        const string outLineSummaryBelowPath = "d:sheetPr/d:outlinePr/@summaryBelow"; 
        public bool OutLineSummaryBelow 
        { 
            get
            {
                return GetXmlNodeBool(outLineSummaryBelowPath);
            }
            set
            {
                SetXmlNode(outLineSummaryBelowPath, value ? "1" : "0");
            }
        }
        const string outLineSummaryRightPath = "d:sheetPr/d:outlinePr/@summaryRight";
        public bool OutLineSummaryRight
        {
            get
            {
                return GetXmlNodeBool(outLineSummaryRightPath);
            }
            set
            {
                SetXmlNode(outLineSummaryRightPath, value ? "1" : "0");
            }
        }
        const string outLineApplyStylePath = "d:sheetPr/d:outlinePr/@applyStyles";
        public bool OutLineApplyStyle
        {
            get
            {
                return GetXmlNodeBool(outLineApplyStylePath);
            }
            set
            {
                SetXmlNode(outLineApplyStylePath, value ? "1" : "0");
            }
        }
        #region WorksheetXml
		/// <summary>
		/// The XML document holding all the worksheet data.
		/// </summary>
		public XmlDocument WorksheetXml
		{
			get
			{
				return (_worksheetXml);
			}
		}

        private void CreateXml()
        {
            _worksheetXml = new XmlDocument();
            _worksheetXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;
            PackagePart packPart = xlPackage.Package.GetPart(WorksheetUri);
            _worksheetXml.Load(packPart.GetStream());

            LoadColumns();
            LoadCells();
            LoadMergeCells();
            LoadHyperLinks();            
        }

        private void LoadColumns()
        {            
            var colList=new List<IRangeID>();
            foreach (XmlNode colNode in _worksheetXml.SelectNodes("//d:cols/d:col", NameSpaceManager))
            {
                int min=int.Parse(colNode.Attributes["min"].Value);
                int max=int.Parse(colNode.Attributes["max"].Value);
                
                int style;
                if (colNode.Attributes["style"]==null || !int.TryParse(colNode.Attributes["style"].Value, out style))
                {
                    style = 0;
                }

                double width = colNode.Attributes["width"] == null ? 0 : double.Parse(colNode.Attributes["width"].Value, _ci); 
                ExcelColumn col = new ExcelColumn(this,min);
                col._columnMax = max;
                col.StyleID = style;
                col.Width = width;

                col.BestFit = colNode.Attributes["bestFit"] != null && colNode.Attributes["bestFit"].Value == "1" ? true : false;
                col.Collapsed = colNode.Attributes["collapsed"] != null && colNode.Attributes["collapsed"].Value == "1" ? true : false;
                col.Phonetic = colNode.Attributes["phonetic"] != null && colNode.Attributes["phonetic"].Value == "1" ? true : false;
                col.OutlineLevel = colNode.Attributes["outlineLevel"] == null ? 0 : int.Parse(colNode.Attributes["outlineLevel"].Value, _ci);
                col.Hidden = colNode.Attributes["hidden"] != null && colNode.Attributes["hidden"].Value == "1" ? true : false;
                colList.Add(col);
            }
            _columns = new RangeCollection(colList);
        }

        private void LoadHyperLinks()
        {
            foreach (XmlElement hlNode in _worksheetXml.SelectNodes("//d:hyperlinks/d:hyperlink", NameSpaceManager))
            {
                int fromRow, fromCol, toRow, toCol;
                ExcelCell.GetRowColFromAddress(hlNode.Attributes["ref"].Value, out fromRow, out fromCol, out toRow, out toCol);
                ulong id = ExcelCell.GetCellID(_sheetID, fromRow, fromCol);
                ExcelCell cell = _cells[id] as ExcelCell;
                if (hlNode.Attributes["r:id"] != null)
                {
                    cell.HyperLinkRId = hlNode.Attributes["r:id"].Value;
                    cell.Hyperlink = Part.GetRelationship(cell.HyperLinkRId).TargetUri;
                }
                else if(hlNode.Attributes["location"]!=null)
                {
                    ExcelHyperLink hl = new ExcelHyperLink(hlNode.GetAttribute("location"), hlNode.GetAttribute("display"));
                    hl.RowSpann = toRow - fromRow;
                    hl.ColSpann = toCol - fromCol;
                    cell.Hyperlink = hl;
                }
            }
        }

        private void LoadCells()
        {
            var cellList=new List<IRangeID>();
            var rowList = new List<IRangeID>();
            var formulaList = new List<IRangeID>(); 
            foreach (XmlNode rowNode in _worksheetXml.SelectNodes("//d:sheetData/d:row", NameSpaceManager))
            {
                int row = Convert.ToInt32(rowNode.Attributes.GetNamedItem("r").Value);
                if (rowNode.Attributes.Count > 2 || (rowNode.Attributes.Count == 2 && rowNode.Attributes.GetNamedItem("spans")!=null))
                {
                    rowList.Add(AddRow(rowNode, row));
                }

                foreach (XmlNode colNode in rowNode.SelectNodes("./d:c", NameSpaceManager))
				{
					ExcelCell cell=new ExcelCell(this, colNode.Attributes["r"].Value);
                    if (colNode.Attributes["t"] != null) cell.DataType = colNode.Attributes["t"].Value;
                    if (colNode.Attributes["s"] != null)
                        cell.StyleID = int.Parse(colNode.Attributes["s"].Value);
                    else
                    {
                        cell.StyleID = 0;
                    }
                    XmlNode f = colNode.SelectSingleNode("./d:f", NameSpaceManager);
                    if (f != null)
                    {
                        XmlNode t = f.Attributes["t"];
                        if (t != null)
                        {
                            if (t.InnerText == "shared")
                            {
                                XmlNode si = f.Attributes["si"];
                                if (si != null)
                                {
                                    cell.SharedFormulaID = int.Parse(si.Value);
                                    if (f.InnerText != "")
                                    {
                                        _sharedFormulas.Add(cell.SharedFormulaID, new Formulas() { Index = cell.SharedFormulaID, Formula = f.InnerText, Address = f.Attributes["ref"].Value, StartRow = cell.Row, StartCol = cell.Column });
                                    }
                                }
                            }
                        }
                        else
                        {
                            cell._formula = f.InnerText;
                        }
                        //Set the value variable (Value property cleans the formla)
                        cell._value = GetValueFromXml(cell, colNode);
                        formulaList.Add(cell);
                    }
                    else
                    {
                        cell._value = GetValueFromXml(cell, colNode);
                    }
                    //_cells.Add(cell.RangeID, cell);
                    cellList.Add(cell);
                }
            }
            _cells = new RangeCollection(cellList);
            _rows = new RangeCollection(rowList);
            _formulaCells = new RangeCollection(formulaList);
        }
        private void LoadMergeCells()
        {
            foreach (XmlNode mergeNode in _worksheetXml.SelectNodes("//d:mergeCells/d:mergeCell", NameSpaceManager))
            {
                string address = mergeNode.Attributes["ref"].Value;
                Cells[address].Merge = true;
            }
        }
        private void UpdateMergedCells()
        {
            var topNode = _worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager);
            if (_mergedCells.Count > 0)
            {
                if (topNode == null)
                {
                    XmlNode parentNode = _worksheetXml.SelectSingleNode("//d:sheetData", NameSpaceManager);
                    topNode = _worksheetXml.CreateElement("mergeCells", ExcelPackage.schemaMain);
                    _worksheetXml.DocumentElement.InsertAfter(topNode, parentNode);

                }
                else
                {
                    topNode.RemoveAll();
                }
                foreach (string address in _mergedCells)
                {
                    XmlElement mergeCell = _worksheetXml.CreateElement("mergeCell", ExcelPackage.schemaMain);
                    mergeCell.SetAttribute("ref", address);
                    topNode.AppendChild(mergeCell);
                }
            }
            else
            {
                if(topNode!=null)  topNode.RemoveAll();
            }
        }
        private ExcelRow AddRow(XmlNode rowNode, int row)
        {
            ExcelRow r = new ExcelRow(this, row);

            r.Collapsed = rowNode.Attributes.GetNamedItem("collapsed") != null && rowNode.Attributes.GetNamedItem("collapsed").Value == "1" ? true: false;
            r.Height=rowNode.Attributes.GetNamedItem("ht")==null ? defaultRowHeight : double.Parse(rowNode.Attributes.GetNamedItem("ht").Value,_ci);
            r.Hidden = rowNode.Attributes.GetNamedItem("hidden") != null && rowNode.Attributes.GetNamedItem("hidden").Value == "1" ? true : false; ;
            r.OutlineLevel = rowNode.Attributes.GetNamedItem("outlineLevel") == null ? 0 : int.Parse(rowNode.Attributes.GetNamedItem("outlineLevel").Value, _ci); ;
            r.Phonetic = rowNode.Attributes.GetNamedItem("ph") != null && rowNode.Attributes.GetNamedItem("ph").Value == "1" ? true : false; ;
            r.StyleID = rowNode.Attributes.GetNamedItem("s") == null ? 0 : int.Parse(rowNode.Attributes.GetNamedItem("s").Value, _ci);
            return r;
        }

        private object GetValueFromXml(ExcelCell cell, XmlNode colNode)
        {
            object value;
            XmlNode vnode = colNode.SelectSingleNode("d:v", NameSpaceManager);
            if (vnode == null) return null;

            string v=vnode.InnerText;
            if (cell.DataType == "s")
            {
                int ix=(int.Parse(v));
                value = xlPackage.Workbook._sharedStringsList[ix].Text;
                cell.IsRichText = xlPackage.Workbook._sharedStringsList[ix].isRichText;
            }
            else if (cell.DataType == "str")
            {
                value = v;
            }
            else
            {
                int n = cell.Style.Numberformat.NumFmtID;
                if ((n >= 14 && n <= 22) || (n >= 45 && n <= 47))
                {
                    double res;
                    if (double.TryParse(v, out res))
                    {
                        value = DateTime.FromOADate(res);
                    }
                    else
                    {
                        value = "";
                    }
                }
                else
                {
                    double d;
                    if (double.TryParse(v, NumberStyles.Any, _ci, out d))
                    {
                        value = d;
                    }
                    else
                    {
                        value = double.NaN;
                    }
                }
            }
            return value;
        }

        private string GetSharedString(int stringID)
        {
            string retValue = null;
            XmlNodeList stringNodes = xlPackage.Workbook.SharedStringsXml.SelectNodes(string.Format("//d:si", stringID), NameSpaceManager);
            XmlNode stringNode = stringNodes[stringID];
            if (stringNode != null)
                retValue = stringNode.InnerText;
            return (retValue);
        }
        #endregion

		#region HeaderFooter
		/// <summary>
		/// A reference to the header and footer class which allows you to 
		/// set the header and footer for all odd, even and first pages of the worksheet
		/// </summary>
		public ExcelHeaderFooter HeaderFooter
		{
			get
			{
				if (_headerFooter == null)
				{
					XmlNode headerFooterNode = WorksheetXml.SelectSingleNode("//d:headerFooter", NameSpaceManager);
					if (headerFooterNode == null)
						headerFooterNode = WorksheetXml.DocumentElement.AppendChild(WorksheetXml.CreateElement("headerFooter", ExcelPackage.schemaMain));
					_headerFooter = new ExcelHeaderFooter((XmlElement)headerFooterNode);
				}
				return (_headerFooter);
			}
		}
		#endregion

        #region "PrinterSettings"
        public ExcelPrinterSettings PrinterSettings
        {
            get
            {
                var ps = new ExcelPrinterSettings(NameSpaceManager, TopNode);
                ps.SchemaNodeOrder = SchemaNodeOrder;
                return ps;
            }
        }
        #endregion
        // TODO: implement freeze pane. 

		#endregion // END Worksheet Public Properties

		#region Worksheet Public Methods
        
        /// <summary>
		/// Provides access to an individual cell within the worksheet.
		/// </summary>
		/// <param name="row">The row number in the worksheet</param>
		/// <param name="col">The column number in the worksheet</param>
		/// <returns></returns>		
        internal ExcelCell Cell(int row, int col)
        {
            ulong cellID=ExcelCell.GetCellID(SheetID, row,col);
            if (!_cells.ContainsKey(cellID))
            {
                _cells.Add(new ExcelCell(this, row, col));
            }
            return _cells[cellID] as ExcelCell;
        }
        /// <summary>
        /// Provide access to a range of cells
        /// </summary>
        public ExcelRange Cells
        {
            get
            {
                return new ExcelRange(this/*,Address*/);
            }
        }
        MergeCellsCollection<string> _mergedCells = new MergeCellsCollection<string>();
        //Dictionary<ulong, ExcelCell> addedMergedCells = new Dictionary<ulong, ExcelCell>();
        public MergeCellsCollection<string> MergedCells
        {
            get
            {
                return _mergedCells;
            }
        }
/*        public ExcelRange Cells(int FromCol, int FromRow, int ToCol, int ToRow)
        {
            return new ExcelRange(this, FromCol, FromRow, ToCol, ToRow);
        }*/
        /// <summary>
		/// Provides access to an individual row within the worksheet so you can set its properties.
		/// </summary>
		/// <param name="row">The row number in the worksheet</param>
		/// <returns></returns>
		public ExcelRow Row(int row)
		{
            ExcelRow r;
            ulong id = ExcelRow.GetRowID(_sheetID, row);
            if (_rows.ContainsKey(id))
            {
                r = _rows[id] as ExcelRow;
            }
            else
            {
                r = new ExcelRow(this, row);
                _rows.Add(r);
            }
            return r;
		}

		/// <summary>
		/// Provides access to an individual column within the worksheet so you can set its properties.
		/// </summary>
		/// <param name="col">The column number in the worksheet</param>
		/// <returns></returns>
		public ExcelColumn Column(int col)
		{
            ExcelColumn column;
            ulong id=ExcelColumn.GetColumnID(_sheetID, col);
            if (_columns.ContainsKey(id))
            {
                column = _columns[id] as ExcelColumn;
                if (column.ColumnMin != column.ColumnMax)
                {
                    int maxCol = column.ColumnMax;
                    column.ColumnMax=col;
                    ExcelColumn copy = Copy(column, col+1);
                    copy.ColumnMax = maxCol;
                }
            }
            else
            {
                foreach (ExcelColumn checkColumn in _columns)
                {
                    if (col > checkColumn.ColumnMin && col <= checkColumn.ColumnMax)
                    {
                        int maxCol = checkColumn.ColumnMax;
                        checkColumn.ColumnMax = col - 1;
                        if (maxCol > col)
                        {
                            ExcelColumn newC = Copy(checkColumn, col + 1);
                            newC.ColumnMax = maxCol;
                        }
                        return Copy(checkColumn, col);                        
                    }
                }
                column = new ExcelColumn(this, col);
                _columns.Add(column);
             }
            return column;
		}

        private ExcelColumn Copy(ExcelColumn c, int col)
        {
            ExcelColumn newC = new ExcelColumn(this, col);
            if (c.StyleName != "")
                newC.StyleName = c.StyleName;
            else
                newC.StyleID = c.StyleID;        
            newC.Width = c.Width;
            newC.Hidden = c.Hidden;
            _columns.Add(newC);
            return newC;
       }

        /// <summary>
        /// Selects a range in the worksheet. The actice cell is the topmost cell.
        /// Make the current worksheet active.
        /// </summary>
        /// <param name="Address">A address range</param>
        public void Select(string Address)
        {
            Select(Address, true);
        }
        /// <summary>
        /// Selects a range in the worksheet. The actice cell is the topmost cell.
        /// </summary>
        /// <param name="Address">A address range</param>
        /// <param name="SelectSheet">Make the sheet active</param>
        public void Select(string Address, bool SelectSheet)
        {
            int fromCol, fromRow, toCol, toRow;
            //Get rows and columns and validate as well
            ExcelCell.GetRowColFromAddress(Address, out fromRow, out fromCol, out toRow, out toCol);

            if (SelectSheet)
            {
                View.TabSelected = true;
            }
            View.SelectedRange = Address;
            View.ActiveCell = ExcelCell.GetAddress(fromRow, fromCol);
        }
        /// <summary>
        /// Inserts conditional formatting for the cell range.
        /// Currently only supports the dataBar style.
        /// </summary>
        /// <param name="startCell"></param>
        /// <param name="endCell"></param>
        /// <param name="color"></param>
        internal void CreateConditionalFormatting(ExcelCell startCell, ExcelCell endCell, string color)
        {
        //    XmlNode formatNode = WorksheetXml.SelectSingleNode("//d:conditionalFormatting", NameSpaceManager);
        //    if (formatNode == null)
        //    {
        //        formatNode = WorksheetXml.CreateElement("conditionalFormatting", ExcelPackage.schemaMain);
        //        XmlNode prevNode = WorksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager);
        //        if (prevNode == null)
        //            prevNode = WorksheetXml.SelectSingleNode("//d:sheetData", NameSpaceManager);
        //        WorksheetXml.DocumentElement.InsertAfter(formatNode, prevNode);
        //    }
        //    XmlAttribute attr = formatNode.Attributes["sqref"];
        //    if (attr == null)
        //    {
        //        attr = WorksheetXml.CreateAttribute("sqref");
        //        formatNode.Attributes.Append(attr);
        //    }
        //    attr.Value = string.Format("{0}:{1}", startCell.CellAddress, endCell.CellAddress);

        //    XmlNode node = formatNode.SelectSingleNode("./d:cfRule", NameSpaceManager);
        //    if (node == null)
        //    {
        //        node = WorksheetXml.CreateElement("cfRule", ExcelPackage.schemaMain);
        //        formatNode.AppendChild(node);
        //    }

        //    attr = node.Attributes["type"];
        //    if (attr == null)
        //    {
        //        attr = WorksheetXml.CreateAttribute("type");
        //        node.Attributes.Append(attr);
        //    }
        //    attr.Value = "dataBar";

        //    attr = node.Attributes["priority"];
        //    if (attr == null)
        //    {
        //        attr = WorksheetXml.CreateAttribute("priority");
        //        node.Attributes.Append(attr);
        //    }
        //    attr.Value = "1";

        //    // the following is poor code, but just an example!!!
        //    XmlNode databar = WorksheetXml.CreateElement("databar", ExcelPackage.schemaMain);
        //    node.AppendChild(databar);

        //    XmlNode child = WorksheetXml.CreateElement("cfvo", ExcelPackage.schemaMain);
        //    databar.AppendChild(child);
        //    attr = WorksheetXml.CreateAttribute("type");
        //    child.Attributes.Append(attr);
        //    attr.Value = "min";
        //    attr = WorksheetXml.CreateAttribute("val");
        //    child.Attributes.Append(attr);
        //    attr.Value = "0";

        //    child = WorksheetXml.CreateElement("cfvo", ExcelPackage.schemaMain);
        //    databar.AppendChild(child);
        //    attr = WorksheetXml.CreateAttribute("type");
        //    child.Attributes.Append(attr);
        //    attr.Value = "max";
        //    attr = WorksheetXml.CreateAttribute("val");
        //    child.Attributes.Append(attr);
        //    attr.Value = "0";

        //    child = WorksheetXml.CreateElement("color", ExcelPackage.schemaMain);
        //    databar.AppendChild(child);
        //    attr = WorksheetXml.CreateAttribute("rgb");
        //    child.Attributes.Append(attr);
        //    attr.Value = color;
            throw(new NotImplementedException("Conditional formatting has been removed for now."));
        }

		#region InsertRow
		/// <summary>
		/// Inserts a new row into the spreadsheet.  Existing rows below the position are 
		/// shifted down.  All formula are updated to take account of the new row.
		/// </summary>
        /// <param name="rowFrom">The position of the new row</param>
        /// <param name="rows">Number of rows to be deleted</param>
		public void InsertRow(int rowFrom, int rows)
		{
            //Insert the new row into the collection
            int index = _cells.InsertRows(ExcelRow.GetRowID(SheetID, rowFrom), rows);
            //List<int> sharedFormulas = new List<int>();
            foreach (ExcelCell cell in _formulaCells)
            {
                if (cell.SharedFormulaID<0)
                {
                    cell.Formula = ExcelCell.UpdateFormulaReferences(cell.Formula, rows, 0, rowFrom, 0);
                }
            }

            FixSharedFormulasRows(rowFrom, rows);
            AddMergedCells(rowFrom, rows);
        }
        /// <summary>
        /// Adds a value to the row of merged cells to fix for inserts or deletes
        /// </summary>
        /// <param name="position"></param>
        /// <param name="rows"></param>
        private void AddMergedCells(int position, int rows)
        {
            for(int i=0;i<_mergedCells.Count;i++)
            {
                int fromRow, toRow, fromCol, toCol;
                ExcelCellBase.GetRowColFromAddress(_mergedCells[i], out fromRow, out  fromCol, out  toRow, out toCol);

                if (fromRow >= position) 
                {
                    fromRow += rows;
                }
                if (toRow >= position)
                {
                    toRow += rows;
                }

                //Set merged prop for cells
                for (int row = fromRow; row <= toRow; row++)
                {
                    for (int col = fromCol; col <= toCol; col++)
                    {
                        Cell(row, col).Merge = true;
                    }
                }

                _mergedCells.List[i] = ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol);
            }
        }
        private void FixSharedFormulasRows(int position, int rows)
        {
            List<Formulas> added=new List<Formulas>();
            List<Formulas> deleted = new List<Formulas>();
            foreach (int id in _sharedFormulas.Keys)
            {
                var f = _sharedFormulas[id];
                int fromCol, fromRow, toCol, toRow;

                ExcelCell.GetRowColFromAddress(f.Address, out fromRow, out fromCol, out toRow, out toCol);
                if (position >= fromRow && position+(Math.Abs(rows)) <= toRow) //Insert/delete is whithin the share formula address
                {
                    if (rows > 0) //Insert
                    {
                        f.Address = ExcelCell.GetAddress(fromRow, fromCol) + ":" + ExcelCell.GetAddress(position - 1, toCol);
                        if (toRow != fromRow)
                        {
                            Formulas newF = new Formulas();
                            newF.StartCol = f.StartCol;
                            newF.StartRow = position + rows;
                            newF.Index = GetMaxShareFunctionIndex();
                            newF.Address = ExcelCell.GetAddress(position + rows, fromCol) + ":" + ExcelCell.GetAddress(toRow + rows, toCol);
                            newF.Formula = ExcelCell.UpdateFormulaReferences(f.Formula, newF.StartRow-f.StartRow, 0, 1, 0); //Recalc the cells positions //ExcelCell.TranslateFromR1C1(Cells[fromRow, fromCol].FormulaR1C1, newF.StartRow, newF.StartCol); //Räkna om formulan
                            Cells[newF.Address].SetSharedFormulaID(newF.Index);
                            added.Add(newF);
                        }
                    }
                    else
                    {
                        if (fromRow - rows < toRow)
                        {
                            f.Address = ExcelCell.GetAddress(fromRow, fromCol, toRow+rows, toCol);
                        }
                        else
                        {
                            f.Address = ExcelCell.GetAddress(fromRow, fromCol) + ":" + ExcelCell.GetAddress(toRow + rows, toCol);
                        }
                    }
                }
                else if (position <= toRow)
                {
                    if (rows > 0) //Insert before shift down
                    {
                        f.StartRow += rows;
                        f.Formula = ExcelCell.UpdateFormulaReferences(f.Formula, rows, 0, position, 0); //Recalc the cells positions
                        f.Address = ExcelCell.GetAddress(fromRow + rows, fromCol) + ":" + ExcelCell.GetAddress(toRow + rows, toCol);
                    }
                    else
                    {
                        if (position <= fromRow && position + Math.Abs(rows) >= toRow)  //Delete the formula 
                        {
                            deleted.Add(f);
                        }
                        else
                        {
                            toRow = toRow + rows < position - 1 ? position - 1 : toRow + rows;
                            if (position <= fromRow)
                            {
                                fromRow = fromRow + rows < position ? position : fromRow + rows;
                            }
                            f.Address = ExcelCell.GetAddress(fromRow, fromCol, toRow, toCol);
                            f.StartRow = fromRow;
                            f.Formula = ExcelCell.UpdateFormulaReferences(f.Formula, rows, 0, position, 0);
                        }
                    }
                }
            }

            //Add new formulas
            foreach(Formulas f in added)
            {
                _sharedFormulas.Add(f.Index, f);
            }
            //Remove formulas
            foreach (Formulas f in deleted)
            {
                _sharedFormulas.Remove(f.Index);
            }
        }
		#endregion

		#region DeleteRow
		/// <summary>
		/// Deletes the specified row from the worksheet.
		/// If shiftOtherRowsUp=true then all formula are updated to take account of the deleted row.
		/// </summary>
		/// <param name="rowFrom">The number of the start row to be deleted</param>
        /// <param name="rows">Number of rows to delete</param>
        /// <param name="shiftOtherRowsUp">Set to true if you want the other rows renumbered so they all move up</param>
		public void DeleteRow(int rowFrom, int rows, bool shiftOtherRowsUp)
		{
            //throw (new Exception("Insert and delete of rows has been removed for now."));

            int index = _cells.DeleteRows(ExcelRow.GetRowID(SheetID, rowFrom), rows);

            foreach (ExcelCell cell in _formulaCells)
            {
                if (cell.SharedFormulaID < 0)
                {
                    cell.Formula = ExcelCell.UpdateFormulaReferences(cell.Formula, rows, 0, rowFrom, 0);
                }
            }
            FixSharedFormulasRows(rowFrom, -rows);
            AddMergedCells(rowFrom, -rows);
        }
		#endregion

		#endregion // END Worksheet Public Methods

		#region Worksheet Private Methods

		#region Worksheet Save
		/// <summary>
		/// Saves the worksheet to the package.  For internal use only.
		/// </summary>
		protected internal void Save()  // Worksheet Save
		{
			#region Delete the printer settings component (if it exists)
			// we also need to delete the relationship from the pageSetup tag
			XmlNode pageSetup = WorksheetXml.SelectSingleNode("//d:pageSetup", NameSpaceManager);
			if (pageSetup != null)
			{
				XmlAttribute attr = (XmlAttribute)pageSetup.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships);
				if (attr != null)
				{
					string relID = attr.Value;
					// first delete the attribute from the XML
					pageSetup.Attributes.Remove(attr);

					// get the URI
					PackageRelationship relPrinterSettings = Part.GetRelationship(relID);
					Uri printerSettingsUri = new Uri("/xl" + relPrinterSettings.TargetUri.ToString().Replace("..", ""), UriKind.Relative);

					// now delete the relationship
					Part.DeleteRelationship(relPrinterSettings.Id);

					// now delete the part from the package
					xlPackage.Package.DeletePart(printerSettingsUri);
				}
			}
			#endregion

			if (_worksheetXml != null)
			{
                
				// save the header & footer (if defined)
				if (_headerFooter != null)
					HeaderFooter.Save();
                // replace the numeric Cell IDs we inserted with AddNumericCellIDs()
				//ReplaceNumericCellIDs();

                if (_cells.Count > 0)
                {
                    this.SetXmlNode("d:dimension/@ref", Dimension);
                }

				// save worksheet to package
				PackagePart partPack = xlPackage.Package.GetPart(WorksheetUri);
				WorksheetXml.Save(Part.GetStream(FileMode.Create, FileAccess.Write));

                xlPackage.WriteDebugFile(WorksheetXml, @"xl\worksheets", "sheet" + SheetID + ".xml");
			}
            
            if (Drawings.UriDrawing!=null)
            {
                PackagePart partPack = Drawings.Part;
				Drawings.DrawingXml.Save(partPack.GetStream(FileMode.Create, FileAccess.Write));
                foreach (ExcelDrawing d in Drawings)
                {
                    if (d is ExcelChart)
                    {
                        ExcelChart c = (ExcelChart)d;
                        c.ChartXml.Save(c.Part.GetStream(FileMode.Create, FileAccess.Write));
                    }
                }   
                //xlPackage.WriteDebugFile(WorksheetXml, @"xl\drawings", "drawing" + SheetID + ".xml");                
            }
		}

        internal void UpdateSheetXml()
        {
            UpdateColumnData();
            UpdateRowCellData();
        }
        /// <summary>
        /// Inserts the cols collection into the XML document
        /// </summary>
        private void UpdateColumnData()
        {
            XmlNode cols = WorksheetXml.SelectSingleNode("//d:cols", NameSpaceManager);
            if (_columns.Count == 0)
            {
                if (cols != null)
                {
                    cols.ParentNode.RemoveChild(cols);
                }
                return;
            }

            if (cols == null)
            {
                XmlNode refNode = WorksheetXml.SelectSingleNode("//d:sheetData", NameSpaceManager);
                cols = WorksheetXml.DocumentElement.InsertBefore(WorksheetXml.CreateElement("cols", ExcelPackage.schemaMain), refNode);
            }
            else
            {
                cols.RemoveAll();
            }
            StringBuilder sbXml = new StringBuilder();

            ExcelColumn prevCol = null;
            foreach (ExcelColumn col in _columns)
            {                
                if (prevCol != null)
                {
                    if(prevCol.ColumnMax != col.ColumnMin-1)
                    {
                        prevCol._columnMax=col.ColumnMin-1;
                    }
                }
                prevCol = col;
            }
            foreach (ExcelColumn col in _columns)
            {
                ExcelStyleCollection<ExcelXfs> cellXfs = xlPackage.Workbook.Styles.CellXfs;

                sbXml.AppendFormat("<col min=\"{0}\" max=\"{1}\"", col.ColumnMin, col.ColumnMax);
                if (col.Hidden == true)
                {
                    //sbXml.Append(" width=\"0\" hidden=\"1\" customWidth=\"1\"");
                    sbXml.Append(" hidden=\"1\"");
                }
                else if (col.BestFit)
                {
                    sbXml.Append(" bestFit=\"1\"");
                }
                if (col.Width != defaultColWidth)
                {
                    sbXml.AppendFormat(_ci, " width=\"{0}\" customWidth=\"1\"", col.Width);
                }
                if (col.OutlineLevel > 0)
                {                    
                    sbXml.AppendFormat(" outlineLevel=\"{0}\" ", col.OutlineLevel);
                    if (col.Collapsed)
                    {
                        if (col.Hidden)
                        {
                            sbXml.Append(" collapsed=\"1\"");
                        }
                        else
                        {
                            sbXml.Append(" collapsed=\"1\" hidden=\"1\""); //Always hidden
                        }
                    }
                }
                if (col.Phonetic)
                {
                    sbXml.Append(" phonetic=\"1\"");
                }
                long styleID = col.StyleID >= 0 ? cellXfs[col.StyleID].newID : col.StyleID;
                if (styleID > 0)
                {
                    sbXml.AppendFormat(" style=\"{0}\"", styleID);
                }
                sbXml.AppendFormat(" />");
            }
            cols.InnerXml = sbXml.ToString();
        }
        /// <summary>
        /// Insert row and cells into the XML document
        /// </summary>
        private void UpdateRowCellData()
        {
            XmlNode top = WorksheetXml.SelectSingleNode("//d:sheetData", NameSpaceManager);
            if (top == null)
            {
                if (_cells.Count == 0)
                    return;
                else
                {
                    top = WorksheetXml.CreateNode(XmlNodeType.Element, "d:sheetData", ExcelPackage.schemaMain);
                    XmlNode parent = _worksheetXml.SelectSingleNode("//d:sheetFormatPr", NameSpaceManager);
                    if (parent == null)
                    {
                        parent = _worksheetXml.SelectSingleNode("//d:sheetViews",NameSpaceManager);                        
                    }
                    _worksheetXml.DocumentElement.InsertAfter(top, parent);
                }
            }
            ExcelStyleCollection<ExcelXfs> cellXfs = xlPackage.Workbook.Styles.CellXfs;
            top.RemoveAll();
            
            List<ulong> hyperLinkCells = new List<ulong>();
            int row = -1;

            StringBuilder sbXml = new StringBuilder();
            var ss = xlPackage.Workbook._sharedStrings;
            foreach (ExcelCell cell in _cells)
            {
                //ExcelCell cell = _cells[cellID];
                long styleID = cell.StyleID >= 0 ? cellXfs[cell.StyleID].newID : cell.StyleID;
                
                //Add the row element if it's a new row
                if (row != cell.Row)
                {
                    if (row != -1) sbXml.Append("</row>");

                    ulong rowID = ExcelRow.GetRowID(SheetID, cell.Row);
                    sbXml.AppendFormat("<row r=\"{0}\" ", cell.Row);
                    if (_rows.ContainsKey(rowID))
                    {
                        ExcelRow currRow = _rows[rowID] as ExcelRow;
                        if (currRow.Hidden == true)
                        {
                            sbXml.Append("ht=\"0\" hidden=\"1\" ");
                        }
                        else if (currRow.Height != defaultRowHeight)
                        {
                            sbXml.AppendFormat(_ci, "ht=\"{0}\" customHeight=\"1\" ", currRow.Height);
                        }   

                        if(currRow.StyleID > 0)
                        {
                            sbXml.AppendFormat("s=\"{0}\" customFormat=\"1\" ", cellXfs[currRow.StyleID].newID);
                        }
                        if (currRow.OutlineLevel > 0)
                        {
                            sbXml.AppendFormat("outlineLevel =\"{0}\" ", currRow.OutlineLevel);
                            if (currRow.Collapsed)
                            {
                                if (currRow.Hidden)
                                {
                                    sbXml.Append(" collapsed=\"1\"");
                                }
                                else
                                {
                                    sbXml.Append(" collapsed=\"1\" hidden=\"1\""); //Always hidden
                                }
                            }
                        }
                        if (currRow.Phonetic)
                        {
                            sbXml.Append(sbXml.Append("ph=\"1\" "));
                        }
                    }
                    sbXml.Append(">");
                    row = cell.Row;
                }
                if (cell.SharedFormulaID >= 0)
                {
                    var f = _sharedFormulas[cell.SharedFormulaID];
                    if (f.StartCol==cell.Column && f.StartRow==cell.Row)
                    {
                        sbXml.AppendFormat("<c r=\"{0}\" s=\"{1}\"><f ref=\"{2}\" t=\"shared\"  si=\"{3}\">{4}</f></c>", cell.CellAddress, styleID < 0 ? 0 : styleID, f.Address, cell.SharedFormulaID, SecurityElement.Escape(f.Formula));
                    }
                    else
                    {
                        sbXml.AppendFormat("<c r=\"{0}\" s=\"{1}\"><f t=\"shared\" si=\"{2}\" /></c>", cell.CellAddress, styleID < 0 ? 0 : styleID, cell.SharedFormulaID);
                    }
                }
                else if (cell.Formula != "")
                {
                    sbXml.AppendFormat("<c r=\"{0}\" s=\"{1}\">", cell.CellAddress, styleID < 0 ? 0 : styleID);
                    sbXml.AppendFormat("<f>{0}</f></c>", SecurityElement.Escape(cell.Formula));
                }
                else
                {
                    if (cell.Value == null)
                    {
                        sbXml.AppendFormat("<c r=\"{0}\" s=\"{1}\" />", cell.CellAddress, styleID < 0 ? 0 : styleID);
                    }
                    else
                    {
                        if ((cell.Value.GetType().IsPrimitive || cell.Value is double || cell.Value is decimal || cell.Value is DateTime) && cell.DataType != "s")
                        {
                            string s;
                            try
                            {
                                if (cell.Value is DateTime)
                                {
                                    s = ((DateTime)cell.Value).ToOADate().ToString(_ci);
                                }
                                else
                                {
                                    s = Convert.ToDecimal(cell.Value, _ci).ToString(_ci);
                                }
                            }

                            catch
                            {
                                s = "0";
                            }
                            sbXml.AppendFormat("<c r=\"{0}\" s=\"{1}\">", cell.CellAddress, styleID < 0 ? 0 : styleID);
                            sbXml.AppendFormat("<v>{0}</v></c>", s);
                        }
                        else
                        {
                            int ix;
                            if (!ss.ContainsKey(cell.Value.ToString()))
                            {
                                ix = ss.Count;
                                ss.Add(cell.Value.ToString(), new ExcelWorkbook.SharedStringItem() { isRichText = cell.IsRichText, pos = ix });
                            }
                            else
                            {
                                ix = ss[cell.Value.ToString()].pos;
                            }
                            sbXml.AppendFormat("<c r=\"{0}\" s=\"{1}\" t=\"s\">", cell.CellAddress, styleID < 0 ? 0 : styleID);
                            sbXml.AppendFormat("<v>{0}</v></c>", ix);
                        }
                        //if (cell.Merge && !addedMergedCells.ContainsKey(cell.CellID))
                        //{
                        //    string address = GetMergeRange(cell, ref addedMergedCells);
                        //    if (address != "")
                        //    {
                        //        mergedCells.Add(address);
                        //    }
                        //}
                    }
                }
                //Update hyperlinks.
                if (cell.Hyperlink != null)
                {
                    hyperLinkCells.Add(cell.CellID);
                }
            }
            if (row != -1) sbXml.Append("</row>");
            top.InnerXml = sbXml.ToString();

            UpdateMergedCells();

            UpdateHyperLinks(hyperLinkCells);
        }

        //private string GetMergeRange(ExcelCell cell, ref Dictionary<ulong, ExcelCell> addedMergedCells)
        //{   
        //    int col = cell.Column+1, row = cell.Row + 1;
        //    ulong cellId = ExcelCell.GetCellID(SheetID, cell.Row, col);
        //    //Get end column
        //    while (_cells.ContainsKey(cellId) && _cells[cellId].Merge)
        //    {                
        //        addedMergedCells.Add(this.Cell(cell.Row, col++).CellID, null);
        //        cellId = ExcelCell.GetCellID(SheetID, cell.Row, col);
        //    }
        //    //Get end row
        //    cellId = ExcelCell.GetCellID(SheetID, row, cell.Column);
        //    while (_cells.ContainsKey(cellId) && _cells[cellId].Merge)
        //    {
        //        addedMergedCells.Add(this.Cell(row++, cell.Column).CellID, null);
        //        cellId = ExcelCell.GetCellID(SheetID, row, cell.Column);
        //    }
        //    if (row == cell.Row + 1 && col == cell.Column + 1)
        //    {
        //        return "";
        //    }
        //    else
        //    {
        //        return string.Format("{0}:{1}", cell.CellAddress, ExcelCell.GetCellAddress(row-1,col-1));
        //    }
        //}
        /// <summary>
        /// Update xml with hyperlinks 
        /// </summary>
        /// <param name="hyperLinkCells">List containing cellid's with hyperlinks</param>
        private void UpdateHyperLinks(List<ulong> hyperLinkCells)
        {
            XmlNode hyperlinkParent = _worksheetXml.SelectSingleNode("//d:hyperlinks", NameSpaceManager);
            if (hyperLinkCells.Count > 0)
            {
                if (hyperlinkParent == null)
                {
                    hyperlinkParent = CreateHyperLinkCollection();
                }
                else
                {
                    //Remove all Relationships
                    foreach(XmlElement e in _worksheetXml.SelectNodes("//d:hyperlink",NameSpaceManager))
                    {
                        string id=e.GetAttribute("id",ExcelPackage.schemaRelationships);
                        if (id != "")
                        {
                            if (Part.RelationshipExists(id))
                            {
                                Part.DeleteRelationship(id);
                            }
                        }
                    }
                    hyperlinkParent.RemoveAll();                    
                }
                Dictionary<string, string> hyps = new Dictionary<string, string>();
                foreach (ulong cellId in hyperLinkCells)
                {
                    ExcelCell cell = _cells[cellId] as ExcelCell;
                    if (cell.Hyperlink is ExcelHyperLink && (cell.Hyperlink as ExcelHyperLink).ReferenceAddress != "")
                    {
                        ExcelHyperLink hl = cell.Hyperlink as ExcelHyperLink;
                        XmlElement node = _worksheetXml.CreateElement("hyperlink", ExcelPackage.schemaMain);
                        node.SetAttribute("ref", Cells[cell.Row, cell.Column, cell.Row+hl.RowSpann, cell.Column+hl.ColSpann].Address);
                        node.SetAttribute("location", ExcelCell.GetFullAddress(Name, hl.ReferenceAddress));
                        node.SetAttribute("display", hl.Display);
                        hyperlinkParent.AppendChild(node);                        
                    }
                    else
                    {
                        string id;
                        if (hyps.ContainsKey(cell.Hyperlink.AbsolutePath))
                        {
                            id = hyps[cell.Hyperlink.AbsolutePath];
                        }
                        else
                        {
                            XmlElement node = _worksheetXml.CreateElement("hyperlink", ExcelPackage.schemaMain);
                            node.SetAttribute("ref", cell.CellAddress);
                            hyperlinkParent.AppendChild(node);

                            XmlAttribute attr = _worksheetXml.CreateAttribute("r", "id", ExcelPackage.schemaRelationships);
                            node.Attributes.Append(attr);
                            PackageRelationship relationship = Part.CreateRelationship(cell.Hyperlink, TargetMode.External, ExcelPackage.schemaHyperlink);
                            attr.Value = relationship.Id;
                            id = relationship.Id;
                        }
                        cell.HyperLinkRId = id;
                    }

                }   
            }
            else if (hyperlinkParent != null)
            {
                _worksheetXml.DocumentElement.RemoveChild(hyperlinkParent);
            }
        }
        /// <summary>
        /// Create the hyperlinks node in the XML
        /// </summary>
        /// <returns></returns>
        private XmlNode CreateHyperLinkCollection()
        {
            XmlElement hl=_worksheetXml.CreateElement("hyperlinks",ExcelPackage.schemaMain);
            XmlNode prevNode = _worksheetXml.SelectSingleNode("//d:conditionalFormatting", NameSpaceManager);
            if (prevNode == null)
            {
                prevNode = _worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager);
                if (prevNode == null)
                {
                    prevNode = _worksheetXml.SelectSingleNode("//d:sheetData", NameSpaceManager);
                }
            }
            return _worksheetXml.DocumentElement.InsertAfter(hl, prevNode);
        }
        public string Dimension
        {
            get
            {
                if (_cells.Count > 0)
                {
                    return ExcelCellBase.GetAddress((_cells[0] as ExcelCell).Row, _minCol, (_cells[_cells.Count - 1] as ExcelCell).Row, _maxCol);
                }
                else
                {
                    return "";
                }

            }
        }
        #region Drawing
        ExcelDrawings _drawings = null;
        public ExcelDrawings Drawings
        {
            get
            {
                if (_drawings == null)
                {
                    _drawings = new ExcelDrawings(xlPackage, this);
                }
                return _drawings;
            }
        }
        #endregion

		/// <summary>
		/// Returns the style ID given a style name.  
		/// The style ID will be created if not found, but only if the style name exists!
		/// </summary>
		/// <param name="StyleName"></param>
		/// <returns></returns>
		protected internal int GetStyleID(string StyleName)
		{
            //// find the named style in the style sheet
            //string searchString = string.Format("//d:cellStyle[@name = '{0}']", StyleName);
            //XmlNode styleNameNode = xlPackage.Workbook.StylesXml.SelectSingleNode(searchString, NameSpaceManager);
            //if (styleNameNode != null)
            //{
            //    string xfId = styleNameNode.Attributes["xfId"].Value;
            //    // look up position of style in the cellXfs 
            //    searchString = string.Format("//d:cellXfs/d:xf[@xfId = '{0}']", xfId);
            //    XmlNode styleNode = xlPackage.Workbook.StylesXml.SelectSingleNode(searchString, NameSpaceManager);
            //    if (styleNode != null)
            //    {
            //        XmlNodeList nodes = styleNode.SelectNodes("preceding-sibling::d:xf", NameSpaceManager);
            //        if (nodes != null)
            //            styleID = nodes.Count;
            //    }
            //}
			ExcelNamedStyleXml namedStyle=null;
            Workbook.Styles.NamedStyles.FindByID(StyleName, ref namedStyle);
            if (namedStyle.XfId == int.MinValue)
            {
                namedStyle.XfId=Workbook.Styles.CellXfs.FindIndexByID(namedStyle.Style.Id);
            }

            //if (namedStyle.XfId!=int.MinValue)
                return namedStyle.XfId;
            //else
            //    return 0;
		}
        public ExcelWorkbook Workbook
        {
            get
            {
                return xlPackage.Workbook;
            }
        }
		#endregion
        #region MergeCells
            //TODO: Implement Medged Cells
        #endregion
        #endregion  // END Worksheet Private Methods

        internal int GetMaxShareFunctionIndex()
        {
            int i = _sharedFormulas.Count + 1;
            while(_sharedFormulas.ContainsKey(i))
            {
                i++;
            }
            return i;
        }
    }  // END class Worksheet
}
